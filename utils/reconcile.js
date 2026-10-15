'use strict';
/**
 * GSTR-2B vs Purchase Register reconciliation — functional style only.
 * Pure helpers + small pipelines. The only mutation is inside local reduce
 * accumulators (never leaks outside a function).
 */
const XLSX = require('xlsx');

/* ───────────────────────────── Config ───────────────────────────── */
const TOLERANCE = 10;               // ₹ tolerance on taxable value for "probable" matches
const MIN_INVOICE_SIMILARITY = 0.4; // 0..1, how alike cleaned invoice numbers must be

/* ───────────────────────── Clean-up utilities ───────────────────── */
const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;

const cleanInvoiceNum = (str) =>
  String(str ?? '').trim().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

const cleanGSTIN = (str) => String(str ?? '').replace(/\s+/g, '').toUpperCase();

const cleanAmount = (val) => {
  if (typeof val === 'number') return Number.isFinite(val) ? round2(val) : 0;
  const text = String(val ?? '').trim();
  if (!text) return 0;
  const negative = /^\(.*\)$/.test(text) || /^-/.test(text);
  const num = parseFloat(text.replace(/[^0-9.]/g, ''));
  return Number.isFinite(num) ? round2(negative ? -num : num) : 0;
};

const normalizeHeader = (h) => String(h ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');

const formatDate = (val) => {
  if (val === undefined || val === null || val === '') return '';
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    return d ? `${String(d.d).padStart(2, '0')}-${String(d.m).padStart(2, '0')}-${d.y}` : String(val);
  }
  return String(val).trim();
};

/* ─────────────────────────── Column mapping ─────────────────────── */
const FIELD_ALIASES = {
  gstin: ['gstin', 'gstinofsupplier', 'suppliergstin', 'vendorgstin', 'partygstin', 'gstno', 'gstinuin', 'gstnumber'],
  invoice: ['invoicenumber', 'invoiceno', 'invoicenum', 'invoice', 'billno', 'billnumber', 'invno', 'documentnumber', 'docno', 'supplierinvoiceno', 'voucherno'],
  taxable: ['taxablevalue', 'taxableamount', 'taxable', 'taxablevaluers', 'basicvalue', 'assessablevalue', 'taxableamt'],
  name: ['tradename', 'legalname', 'tradelegalname', 'suppliername', 'vendorname', 'partyname', 'name', 'particulars', 'supplier'],
  date: ['invoicedate', 'billdate', 'date', 'documentdate', 'invdate'],
  igst: ['igst', 'integratedtax', 'integratedtaxamount', 'igstamount'],
  cgst: ['cgst', 'centraltax', 'centraltaxamount', 'cgstamount'],
  sgst: ['sgst', 'sgstutgst', 'stateuttax', 'statetax', 'stateuttaxamount', 'sgstamount', 'utgst'],
  totalTax: ['totaltax', 'totalgst', 'gstamount', 'taxamount', 'gst'],
};
const REQUIRED_FIELDS = ['gstin', 'invoice', 'taxable'];
const FIELD_LABELS = { gstin: 'GSTIN', invoice: 'Invoice Number', taxable: 'Taxable Value' };

const buildColumnMap = (headerRow) => {
  const normalized = headerRow.map(normalizeHeader);
  return Object.fromEntries(
    Object.entries(FIELD_ALIASES).map(([field, aliases]) => [
      field,
      // alias order = priority (e.g. "Supplier Invoice No" beats "Voucher No")
      aliases.map((a) => normalized.indexOf(a)).find((i) => i >= 0) ?? -1,
    ])
  );
};

const scoreColumnMap = (map) => REQUIRED_FIELDS.filter((f) => map[f] >= 0).length;

const isFilled = (c) => String(c ?? '').trim() !== '';

/** Overlay two consecutive header rows (official GSTR-2B has a 2-row merged header). */
const mergeHeaderRows = (upper = [], lower = []) =>
  Array.from({ length: Math.max(upper.length, lower.length) }, (_, k) =>
    isFilled(lower[k]) ? lower[k] : upper[k]
  );

/**
 * Looks through the first 25 rows for the best header, trying each row on its own and
 * merged with the row below it. `index` is the LAST header row, so data starts at index + 1.
 */
const findHeaderRow = (rows) =>
  rows
    .slice(0, 25)
    .flatMap((row, i) => [
      { index: i, headers: row },
      { index: i + 1, headers: mergeHeaderRows(row, rows[i + 1]) },
    ])
    .reduce(
      (best, cand) => {
        const map = buildColumnMap(cand.headers);
        const score = scoreColumnMap(map);
        return score > best.score ? { index: cand.index, score, map, headers: cand.headers } : best;
      },
      { index: -1, score: -1, map: {}, headers: [] }
    );

/* ─────────────────────────── Validation ─────────────────────────── */
const TAX_FIELDS = ['igst', 'cgst', 'sgst', 'totalTax'];

/** True when a raw cell is blank, a dash placeholder, or a clean number (₹, commas, (neg) allowed). */
const isValidAmountCell = (val) => {
  if (val === undefined || val === null) return true;
  if (typeof val === 'number') return Number.isFinite(val);
  const t = String(val).replace(/[₹,\s]|rs\.?|inr/gi, '');
  return t === '' || /^[-—–]+$/.test(t) || /^\(?-?(\d+\.?\d*|\.\d+)\)?$/.test(t);
};

/** Returns every problem found in a raw-fields object. `key` maps to an input in the fix form. */
const validateFields = (f) => {
  const issues = [
    !cleanGSTIN(f.gstin) && { reason: 'Missing GSTIN', key: 'gstin' },
    !cleanInvoiceNum(f.invoice) && { reason: 'Missing Invoice Number', key: 'invoice' },
    !isValidAmountCell(f.taxable) && { reason: 'Invalid Taxable Value', key: 'taxable' },
    TAX_FIELDS.some((k) => !isValidAmountCell(f[k])) && { reason: 'Invalid Tax Amount', key: 'tax' },
  ].filter(Boolean);
  return { issues, reason: issues.map((i) => i.reason).join('; ') };
};

/* ─────────────────────────── Parsing ────────────────────────────── */
const isBlankRow = (row) => row.every((c) => String(c ?? '').trim() === '');

/** Pull the raw (uncleaned) value of every known field out of one sheet row. */
const rowToFields = (columnMap, row) =>
  Object.fromEntries(Object.keys(FIELD_ALIASES).map((f) => [f, columnMap[f] >= 0 ? row[columnMap[f]] : undefined]));

/** Raw fields -> normalized invoice record. */
const fieldsToRecord = (f, excelRow) => {
  const igst = cleanAmount(f.igst);
  const cgst = cleanAmount(f.cgst);
  const sgst = cleanAmount(f.sgst);
  const split = round2(igst + cgst + sgst);
  return {
    row: excelRow,
    gstin: cleanGSTIN(f.gstin),
    invoice: cleanInvoiceNum(f.invoice),
    invoiceRaw: String(f.invoice ?? '').trim(),
    name: String(f.name ?? '').trim(),
    date: formatDate(f.date),
    taxable: cleanAmount(f.taxable),
    igst, cgst, sgst,
    tax: split || cleanAmount(f.totalTax),
  };
};

/** Header-labelled view of a row (filled cells only) for display in the Skipped Rows tab. */
const toRawData = (headers, row) =>
  row.reduce((acc, cell, k) => {
    if (!isFilled(cell)) return acc;
    const label = String(headers[k] ?? '').trim() || `Column ${XLSX.utils.encode_col(k)}`;
    const key = label in acc ? `${label} (${XLSX.utils.encode_col(k)})` : label;
    return { ...acc, [key]: cell };
  }, {});

/**
 * Parse one uploaded file buffer. Returns {ok, records, skippedRows, sheet} or {ok:false, error}.
 * Rows that fail validation are NEVER dropped silently — they come back in `skippedRows`
 * as { id, source, rowIndex, rawData, fields, reasonForSkipping, badFields }.
 */
const parseFile = (file, label, source = label === 'GSTR-2B' ? 'portal' : 'books') => {
  let workbook;
  try {
    workbook = XLSX.read(file.buffer, { type: 'buffer', raw: true });
  } catch (e) {
    return { ok: false, error: `Couldn't read the ${label} file "${file.originalname}". Make sure it is a valid .xlsx, .xls or .csv file.` };
  }

  // Keep blank rows while reading so array position -> true Excel row number; then drop them for header detection.
  const candidates = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name];
    const startRow = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']).s.r : 0;
    const filled = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: true })
      .map((row, i) => ({ row, excelRow: startRow + i + 1 }))
      .filter(({ row }) => !isBlankRow(row));
    const rows = filled.map((x) => x.row);
    return { name, filled, rows, header: findHeaderRow(rows) };
  }).sort((a, b) =>
    b.header.score - a.header.score ||
    Number(/^b2b$/i.test(b.name)) - Number(/^b2b$/i.test(a.name)) ||
    b.rows.length - a.rows.length
  );

  const best = candidates[0];
  if (!best || best.rows.length === 0) {
    return { ok: false, error: `The ${label} file "${file.originalname}" looks empty. Please upload a file that has data rows.` };
  }

  const { index, score, map, headers } = best.header;
  if (score < REQUIRED_FIELDS.length) {
    const missing = REQUIRED_FIELDS.filter((f) => map[f] < 0).map((f) => FIELD_LABELS[f]);
    const seen = (headers || best.rows[0] || []).filter(Boolean).slice(0, 10).join(', ') || 'none found';
    return {
      ok: false,
      error: `The ${label} file is missing column(s): ${missing.join(', ')}. Headers we found: ${seen}. Rename your columns (e.g. "GSTIN", "Invoice Number", "Taxable Value") and try again.`,
    };
  }

  const parsed = best.filled.slice(index + 1).map(({ row, excelRow }) => {
    const fields = rowToFields(map, row);
    return { row, excelRow, fields, check: validateFields(fields) };
  });

  const prefix = label[0];
  const records = parsed
    .filter((p) => !p.check.issues.length)
    .map((p, id) => ({ ...fieldsToRecord(p.fields, p.excelRow), id: `${prefix}${id}` }));

  const skippedRows = parsed
    .filter((p) => p.check.issues.length)
    .map((p) => ({
      id: `${source}-${p.excelRow}`,
      source,
      sourceLabel: label,
      rowIndex: p.excelRow,
      rawData: toRawData(headers, p.row),
      fields: p.fields,
      reasonForSkipping: p.check.reason,
      badFields: p.check.issues.map((i) => i.key),
    }));

  if (records.length === 0) {
    return { ok: false, error: `No usable rows found in the ${label} file. Each row needs a GSTIN and an invoice number.` };
  }
  return { ok: true, records, skippedRows, sheet: best.name };
};

const buildWarnings = (skippedRows) =>
  Object.entries(
    skippedRows.reduce((acc, s) => ({ ...acc, [s.sourceLabel]: [...(acc[s.sourceLabel] || []), s] }), {})
  ).map(([label, rows]) =>
    `${label}: skipped ${rows.length} row(s) that failed validation (Excel rows ${rows.slice(0, 8).map((r) => r.rowIndex).join(', ')}${rows.length > 8 ? '…' : ''}).`
  );

/** Re-process one skipped row with user-supplied fixes. Pure: returns new `inputs` or an error. */
const reprocessRow = (inputs, rowId, fix = {}) => {
  const target = inputs.skippedRows.find((s) => s.id === rowId);
  if (!target) return { ok: false, error: 'That row was not found (it may already have been re-processed).' };
  const given = (k) => String(fix[k] ?? '').trim();
  const fields = {
    ...target.fields,
    ...(given('gstin') && { gstin: given('gstin') }),
    ...(given('invoice') && { invoice: given('invoice') }),
    ...(given('taxable') && { taxable: given('taxable') }),
    ...(given('tax') && { igst: '', cgst: '', sgst: '', totalTax: given('tax') }),
  };
  const check = validateFields(fields);
  if (check.issues.length) return { ok: false, error: `Still invalid: ${check.reason}.` };
  const prefix = target.source === 'books' ? 'P' : 'G';
  const record = { ...fieldsToRecord(fields, target.rowIndex), id: `${prefix}r${target.rowIndex}`, reprocessed: true };
  return {
    ok: true,
    inputs: {
      ...inputs,
      [target.source]: [...inputs[target.source], record],
      skippedRows: inputs.skippedRows.filter((s) => s.id !== rowId),
    },
  };
};

/* ─────────────────────────── Matching ───────────────────────────── */
const levenshtein = (a, b) =>
  [...a].reduce(
    (prev, ca, i) =>
      [...b].reduce(
        (row, cb, j) => [...row, Math.min(prev[j + 1] + 1, row[j] + 1, prev[j] + (ca === cb ? 0 : 1))],
        [i + 1]
      ),
    Array.from({ length: b.length + 1 }, (_, i) => i)
  ).at(-1);

const similarity = (a, b) => {
  const longest = Math.max(a.length, b.length);
  return longest === 0 ? 1 : 1 - levenshtein(a, b) / longest;
};

const groupBy = (items, keyFn) =>
  items.reduce((map, item) => {
    const key = keyFn(item);
    (map.get(key) || map.set(key, []).get(key)).push(item);
    return map;
  }, new Map());

const sumBy = (items, fn) => round2(items.reduce((total, item) => total + fn(item), 0));

const pairKey = (r) => `${r.gstin}|${r.invoice}`;
const exactKey = (r) => `${pairKey(r)}|${r.taxable.toFixed(2)}`;

/**
 * Generic matching stage. `pick(book, index, used)` returns {rec, meta} or null.
 * Returns matches + whatever is left unmatched on both sides.
 */
const runStage = (books, portal, indexKeyFn, pick) => {
  const index = groupBy(portal, indexKeyFn);
  const acc = books.reduce(
    (state, book) => {
      const hit = pick(book, index, state.used);
      if (!hit) { state.restBooks.push(book); return state; }
      state.used.add(hit.rec.id);
      state.matches.push({ books: book, portal: hit.rec, ...hit.meta });
      return state;
    },
    { matches: [], restBooks: [], used: new Set() }
  );
  return {
    matches: acc.matches,
    restBooks: acc.restBooks,
    restPortal: portal.filter((p) => !acc.used.has(p.id)),
  };
};

const firstUnused = (list = [], used) => list.find((r) => !used.has(r.id));

const exactStage = (books, portal) =>
  runStage(books, portal, exactKey, (book, index, used) => {
    const rec = firstUnused(index.get(exactKey(book)), used);
    return rec ? { rec, meta: { reason: 'Exact match', similarity: 1 } } : null;
  });

const sameInvoiceStage = (books, portal) =>
  runStage(books, portal, pairKey, (book, index, used) => {
    const rec = firstUnused(index.get(pairKey(book)), used);
    if (!rec) return null;
    const diff = round2(book.taxable - rec.taxable);
    return { rec, meta: { reason: `Taxable value differs by ₹${Math.abs(diff).toFixed(2)}`, similarity: 1 } };
  });

const probableStage = (books, portal) =>
  runStage(books, portal, (r) => r.gstin, (book, index, used) => {
    const best = (index.get(book.gstin) || [])
      .filter((p) => !used.has(p.id) && Math.abs(book.taxable - p.taxable) <= TOLERANCE)
      .map((p) => ({ rec: p, sim: similarity(book.invoice, p.invoice), gap: Math.abs(book.taxable - p.taxable) }))
      .filter((c) => c.sim >= MIN_INVOICE_SIMILARITY)
      .reduce((top, c) => (!top || c.sim > top.sim || (c.sim === top.sim && c.gap < top.gap) ? c : top), null);
    return best
      ? { rec: best.rec, meta: { reason: 'Invoice number differs (probable typo)', similarity: round2(best.sim) } }
      : null;
  });

/** Core pipeline: exact → same-invoice-value-diff → probable → leftovers. */
const matchRecords = (books, portal) => {
  const s1 = exactStage(books, portal);
  const s2 = sameInvoiceStage(s1.restBooks, s1.restPortal);
  const s3 = probableStage(s2.restBooks, s2.restPortal);
  return {
    exact: s1.matches,
    mismatches: [...s2.matches, ...s3.matches],
    missingInPortal: s3.restBooks,
    missingInBooks: s3.restPortal,
  };
};

/* ─────────────────────────── Summary ────────────────────────────── */
const bucketStats = (records) => ({
  count: records.length,
  taxable: sumBy(records, (r) => r.taxable),
  tax: sumBy(records, (r) => r.tax),
});

const buildSummary = (booksCount, portalCount, r) => {
  const excessInMismatch = sumBy(r.mismatches, (m) => Math.max(0, m.books.tax - m.portal.tax));
  return {
    booksCount,
    portalCount,
    exact: bucketStats(r.exact.map((m) => m.books)),
    mismatches: bucketStats(r.mismatches.map((m) => m.books)),
    missingInPortal: bucketStats(r.missingInPortal),
    missingInBooks: bucketStats(r.missingInBooks),
    discrepancyCount: r.mismatches.length + r.missingInPortal.length + r.missingInBooks.length,
    itcAtRisk: round2(sumBy(r.missingInPortal, (x) => x.tax) + excessInMismatch),
    matchRate: booksCount ? Math.round((r.exact.length / booksCount) * 100) : 0,
  };
};

/* ───────────────────── Follow-up links (Missing in GSTR-2B) ───────────────────── */
const formatINR = (n) =>
  Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const reminderMessage = (r) =>
  `Hi ${r.name || 'Sir/Madam'}, your invoice ${r.invoiceRaw} dated ${r.date || 'N/A'} for ₹${formatINR(r.tax)} tax amount ` +
  'is currently missing from GSTR-2B. Please upload this in your GSTR-1 return at the earliest so we can process your pending payment.';

const reminderSubject = (r) => `Urgent: Missing Invoice ${r.invoiceRaw} in GSTR-2B`;
const whatsappLink = (r) => `https://wa.me/?text=${encodeURIComponent(reminderMessage(r))}`;
const mailtoLink = (r) =>
  `mailto:?subject=${encodeURIComponent(reminderSubject(r))}&body=${encodeURIComponent(reminderMessage(r))}`;

const withActionLinks = (r) => ({ ...r, links: { whatsapp: whatsappLink(r), email: mailtoLink(r) } });

/* ───────────────────── Periods & unreconciled ITC ledger ───────────────────── */
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const isPeriod = (p) => /^\d{4}-(0[1-9]|1[0-2])$/.test(String(p ?? ''));
/** '2026-09' -> 'Sep-2026' */
const periodLabel = (p) => (isPeriod(p) ? `${MONTHS[Number(p.slice(5)) - 1]}-${p.slice(0, 4)}` : String(p || ''));

/** Pull 'YYYY-MM' out of dd-mm-yyyy / dd/mm/yyyy / yyyy-mm-dd strings. */
const dateToPeriod = (s) => {
  const dmy = String(s).match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  const ymd = String(s).match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  const [y, m] = dmy ? [dmy[3], dmy[2]] : ymd ? [ymd[1], ymd[2]] : [null, null];
  const p = y ? `${y}-${String(m).padStart(2, '0')}` : '';
  return isPeriod(p) ? p : '';
};

/** Most common invoice month in the GSTR-2B records (used when the user doesn't pick a period). */
const derivePeriod = (records, today = new Date()) => {
  const counts = [...groupBy(records.map((r) => dateToPeriod(r.date)).filter(Boolean), (p) => p)]
    .map(([p, list]) => [p, list.length])
    .sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? 1 : -1));
  return counts[0]?.[0] || `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
};

const emptyLedger = () => ({ pending: [], claimed: [] });

const toLedgerEntry = (period, now) => (r) => ({
  id: `${period}|${pairKey(r)}`,
  period,
  periodLabel: periodLabel(period),
  gstin: r.gstin, invoice: r.invoice, invoiceRaw: r.invoiceRaw, name: r.name, date: r.date,
  taxable: r.taxable, igst: r.igst, cgst: r.cgst, sgst: r.sgst, tax: r.tax,
  savedAt: now,
});

/**
 * Rollover check: which invoices flagged "missing" in an EARLIER month show up in this month's
 * GSTR-2B? `portalLeft` should be the portal records not matched to this month's books.
 * Entries already claimed in this same period are included so re-opening a report is idempotent.
 */
const rolloverMatch = (ledger, portalLeft, period) => {
  const candidates = [...ledger.pending, ...ledger.claimed.filter((e) => e.claimedIn === period)]
    .filter((e) => e.period < period);
  const index = groupBy(portalLeft, pairKey);
  return candidates.reduce(
    (state, entry) => {
      const rec = firstUnused(index.get(pairKey(entry)), state.used);
      if (!rec) return state;
      state.used.add(rec.id);
      const diff = round2(rec.taxable - entry.taxable);
      state.hits.push({
        ledger: entry,
        portal: rec,
        taxableDiff: diff,
        message: diff === 0
          ? `Invoice ${entry.invoiceRaw} from ${entry.periodLabel} was filed late by supplier in ${periodLabel(period)} GSTR-2B. Safe to claim now!`
          : `Invoice ${entry.invoiceRaw} from ${entry.periodLabel} now appears in ${periodLabel(period)} GSTR-2B, but its taxable value differs by ₹${formatINR(Math.abs(diff))} from your books. Verify before claiming.`,
      });
      return state;
    },
    { hits: [], used: new Set() }
  ).hits;
};

/** Pure ledger transition after a run: clear matched entries, replace this period's pending list. */
const updateLedger = (ledger, { period, missing, matchedIds, now = new Date().toISOString() }) => {
  const matched = new Set(matchedIds);
  const alreadyClaimed = new Set(ledger.claimed.map((e) => e.id));
  const newlyClaimed = ledger.pending
    .filter((e) => matched.has(e.id))
    .map((e) => ({ ...e, claimedIn: period, claimedAt: now }));
  const fresh = missing.map(toLedgerEntry(period, now)).filter((e) => !alreadyClaimed.has(e.id) && !matched.has(e.id));
  return {
    pending: [...ledger.pending.filter((e) => !matched.has(e.id) && e.period !== period), ...fresh],
    claimed: [...ledger.claimed, ...newlyClaimed],
  };
};

/* ─────────────────────────── Entry point ────────────────────────── */
const reconcileRecords = (books, portal) => {
  const results = matchRecords(books, portal);
  return { results, summary: buildSummary(books.length, portal.length, results) };
};

/**
 * Full pipeline on parsed inputs: this month's match + rollover against the saved ledger.
 * Pure — returns the new ledger as `ledger` for the caller to persist.
 */
const buildOutcome = (inputs, ledger = emptyLedger(), now = new Date().toISOString()) => {
  const { books, portal, skippedRows, period } = inputs;
  const base = reconcileRecords(books, portal);
  const late = rolloverMatch(ledger, base.results.missingInBooks, period);
  const lateIds = new Set(late.map((l) => l.portal.id));

  const results = {
    ...base.results,
    missingInPortal: base.results.missingInPortal.map(withActionLinks),
    missingInBooks: base.results.missingInBooks.filter((p) => !lateIds.has(p.id)),
    lateClaimed: late,
  };
  const summary = {
    ...buildSummary(books.length, portal.length, results),
    lateClaimed: { count: late.length, tax: sumBy(late, (l) => l.portal.tax) },
    skippedCount: skippedRows.length,
  };
  return {
    results,
    summary,
    skippedRows,
    period,
    periodLabel: periodLabel(period),
    warnings: buildWarnings(skippedRows),
    ledger: updateLedger(ledger, {
      period, now,
      missing: base.results.missingInPortal,
      matchedIds: late.map((l) => l.ledger.id),
    }),
  };
};

/** Takes the two multer file objects; opts = { period: 'YYYY-MM', ledger }. */
const processUpload = (booksFile, portalFile, opts = {}) => {
  const books = parseFile(booksFile, 'Purchase Register', 'books');
  if (!books.ok) return books;
  const portal = parseFile(portalFile, 'GSTR-2B', 'portal');
  if (!portal.ok) return portal;
  const inputs = {
    books: books.records,
    portal: portal.records,
    skippedRows: [...books.skippedRows, ...portal.skippedRows],
    period: isPeriod(opts.period) ? opts.period : derivePeriod(portal.records),
  };
  return { ok: true, inputs, ...buildOutcome(inputs, opts.ledger || emptyLedger()) };
};

module.exports = {
  TOLERANCE, MIN_INVOICE_SIMILARITY,
  cleanInvoiceNum, cleanGSTIN, cleanAmount,
  validateFields, parseFile, reprocessRow,
  matchRecords, reconcileRecords, buildOutcome, processUpload,
  reminderMessage, reminderSubject, whatsappLink, mailtoLink,
  isPeriod, periodLabel, derivePeriod, emptyLedger, rolloverMatch, updateLedger,
};
