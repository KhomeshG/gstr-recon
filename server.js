'use strict';
const path = require('path');
const express = require('express');
const multer = require('multer');

const { processUpload, buildOutcome, reprocessRow, periodLabel } = require('./utils/reconcile');
const { buildExportBuffer, buildNonCompliantBuffer } = require('./utils/exporter');
const { createStore } = require('./utils/store');
const { createLedgerFile } = require('./utils/ledger');
const samples = require('./utils/sample');

const PORT = process.env.PORT || 3000;
const ALLOWED_EXT = ['.xlsx', '.xls', '.csv'];
const MAX_MB = 15;

const store = createStore();
const ledgerFile = createLedgerFile();
const app = express();
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const TAB_KEYS = ['exact', 'mismatch', 'portal', 'books', 'late', 'skipped'];

app.use(express.json());
app.use(express.urlencoded({ extended: false }));

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.locals.inr = (n) =>
  Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
app.locals.periodLabel = periodLabel;

/* ── Upload handling ───────────────────────────────────────────── */
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_MB * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = ALLOWED_EXT.includes(path.extname(file.originalname).toLowerCase());
    cb(ok ? null : new multer.MulterError('LIMIT_UNEXPECTED_FILE', file.fieldname), ok);
  },
}).fields([{ name: 'books', maxCount: 1 }, { name: 'portal', maxCount: 1 }]);

const uploadErrorMessage = (err) =>
  err.code === 'LIMIT_FILE_SIZE' ? `One of the files is larger than ${MAX_MB} MB.`
  : err.code === 'LIMIT_UNEXPECTED_FILE' ? 'Only .xlsx, .xls and .csv files are supported.'
  : 'The upload failed. Please try again.';

const withUpload = (req, res, next) =>
  upload(req, res, (err) =>
    err ? res.status(400).render('index', { error: uploadErrorMessage(err) }) : next()
  );

const missingFilesMessage = (files = {}) => {
  const missing = [
    !files.books?.[0] && 'Purchase Register (Books)',
    !files.portal?.[0] && 'GSTR-2B Statement (Portal)',
  ].filter(Boolean);
  return missing.length ? `Please choose a file for: ${missing.join(' and ')}.` : null;
};

/* ── Routes ────────────────────────────────────────────────────── */
app.get('/', (req, res) => res.render('index', { error: null }));

/** Ledger is persisted separately from the (expiring) in-memory report. */
const persistOutcome = ({ ledger, ...session }) => {
  ledgerFile.save(ledger);
  return session;
};

const expired = (res) =>
  res.status(404).render('index', { error: 'That report has expired. Please upload your files again.' });

app.post('/reconcile', withUpload, (req, res) => {
  const problem = missingFilesMessage(req.files);
  if (problem) return res.status(400).render('index', { error: problem });

  try {
    const outcome = processUpload(req.files.books[0], req.files.portal[0], {
      period: req.body?.period,
      ledger: ledgerFile.load(),
    });
    if (!outcome.ok) return res.status(422).render('index', { error: outcome.error });

    const session = persistOutcome(outcome);
    const id = store.save({
      ...session,
      files: { books: req.files.books[0].originalname, portal: req.files.portal[0].originalname },
    });
    return res.redirect(303, `/results/${id}`); // PRG: refresh/back never re-submits the upload
  } catch (e) {
    console.error(e);
    return res.status(500).render('index', { error: 'Something went wrong while reconciling. Please check your files and try again.' });
  }
});

app.get('/results/:id', (req, res) => {
  const data = store.get(req.params.id);
  if (!data) return expired(res);
  const tab = TAB_KEYS.includes(req.query.tab) ? req.query.tab : '';
  return res.render('results', { ...data, id: req.params.id, tab });
});

/** Re-process one skipped row with the GSTIN / invoice / amount the user typed in. */
app.post('/reprocess/:id', (req, res) => {
  const data = store.get(req.params.id);
  if (!data) return res.status(404).json({ ok: false, error: 'This report has expired. Please upload your files again.' });

  const fixed = reprocessRow(data.inputs, String(req.body?.rowId ?? ''), req.body);
  if (!fixed.ok) return res.status(422).json(fixed);

  try {
    const session = persistOutcome(buildOutcome(fixed.inputs, ledgerFile.load()));
    store.set(req.params.id, { ...data, ...session, inputs: fixed.inputs });
    return res.json({ ok: true, remaining: session.skippedRows.length });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: 'Could not re-process that row.' });
  }
});

app.get('/export/:id', (req, res) => {
  const data = store.get(req.params.id);
  if (!data) return expired(res);
  res.set({ 'Content-Type': XLSX_MIME, 'Content-Disposition': 'attachment; filename="GSTR2B-Reconciliation.xlsx"' });
  return res.send(buildExportBuffer(data));
});

/** Follow-up sheet for suppliers whose invoices are missing from GSTR-2B. Usage: /export-non-compliant?id=<reportId> */
app.get('/export-non-compliant', (req, res) => {
  const data = store.get(String(req.query.id ?? ''));
  if (!data) return expired(res);
  res.set({
    'Content-Type': XLSX_MIME,
    'Content-Disposition': `attachment; filename="Non-Compliant-Suppliers-${data.periodLabel}.xlsx"`,
  });
  return res.send(buildNonCompliantBuffer(data.results.missingInPortal));
});

/* ── Saved history (unreconciled ITC ledger) ───────────────────── */
app.get('/api/ledger', (req, res) => {
  const { pending, claimed } = ledgerFile.load();
  res.json({
    pendingCount: pending.length,
    pendingTax: Math.round(pending.reduce((s, e) => s + e.tax, 0) * 100) / 100,
    pending,
    claimed,
  });
});

app.delete('/api/ledger/:entryId', (req, res) => {
  const ledger = ledgerFile.load();
  const next = { ...ledger, pending: ledger.pending.filter((e) => e.id !== req.params.entryId) };
  if (next.pending.length === ledger.pending.length) return res.status(404).json({ ok: false, error: 'Entry not found.' });
  ledgerFile.save(next);
  return res.json({ ok: true });
});

app.get('/ledger', (req, res) => res.render('ledger', { ledger: ledgerFile.load() }));

app.get('/sample/:name', (req, res) => {
  const csv = samples[req.params.name];
  if (!csv) return res.status(404).send('Not found');
  res.set({ 'Content-Type': 'text/csv; charset=utf-8', 'Content-Disposition': `attachment; filename="${req.params.name}"` });
  return res.send(csv);
});

app.use((req, res) => res.status(404).render('index', { error: 'Page not found.' }));

app.listen(PORT, () => console.log(`GSTR-2B Reconciler running → http://localhost:${PORT}`));
