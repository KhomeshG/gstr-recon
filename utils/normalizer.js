'use strict';
/**
 * Zero-prep data normalization pipeline — pure functions only.
 *
 *   raw .xlsx/.xls/.csv buffer
 *     → detect the real header row (skipping banners / address blocks)
 *     → fuzzy-map columns to canonical fields (GSTIN, invoice, taxable, …)
 *     → drop footer / summary / repeated-header / key-less junk rows
 *     → scrub every cell (invoice numbers, ₹ amounts, dates)
 *
 * If a mandatory field can't be mapped with high confidence the pipeline does NOT
 * crash: it returns { needsManualMapping: true, request } so the UI can ask the user.
 */
const XLSX = require('xlsx');

/* ───────────────────────────── Config ───────────────────────────── */
const HEADER_SCAN_ROWS = 20;     // how many top rows we search for the header
const PREVIEW_ROWS = 30;         // rows shown in the manual-mapping preview
const PREVIEW_COLS = 40;
const CONFIDENT = 0.7;           // mandatory fields must map at least this well
const MIN_USABLE = 0.5;          // optional fields are used from this confidence up
const SIDE_LABEL = { books: 'Purchase Register', portal: 'GSTR-2B' };

/* ───────────────────────── Small helpers ────────────────────────── */
const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const isFilled = (c) => String(c ?? '').trim() !== '';
const isBlankRow = (row) => row.every((c) => !isFilled(c));
const normalizeHeader = (h) => String(h ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
const plural = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;

const colLetter = (i) => {
  const go = (n) => (n < 0 ? '' : go(Math.floor(n / 26) - 1) + String.fromCharCode(65 + (n % 26)));
  return go(i);
};

/** Edit distance where swapping two adjacent letters ("Invioce") costs 1, not 2. */
const editDistance = (a, b) => {
  const A = [...a];
  const B = [...b];
  const table = A.reduce(
    (rows, ca, i) => {
      const prev = rows[i];
      const prev2 = rows[i - 1];
      const row = B.reduce((r, cb, j) => {
        const base = Math.min(prev[j + 1] + 1, r[j] + 1, prev[j] + (ca === cb ? 0 : 1));
        const swap = i > 0 && j > 0 && ca === B[j - 1] && A[i - 1] === cb ? prev2[j - 1] + 1 : Infinity;
        return [...r, Math.min(base, swap)];
      }, [i + 1]);
      return [...rows, row];
    },
    [Array.from({ length: B.length + 1 }, (_, j) => j)]
  );
  return table.at(-1).at(-1);
};

const similarity = (a, b) => {
  const longest = Math.max(a.length, b.length);
  return longest === 0 ? 1 : 1 - editDistance(a, b) / longest;
};

/* ═══════════════════════ 1. Cell scrubbing ═══════════════════════ */

/* ── Invoice numbers ─────────────────────────────────────────────── */
const SCIENTIFIC = /^[+-]?\d+(\.\d+)?e[+-]?\d+$/i;

/** 101000 (number) or 1.01E+05 (text) → "101000". Never prints exponent notation. */
const expandNumber = (n) =>
  Number.isInteger(n) && Math.abs(n) < 1e21 ? BigInt(n).toString() : String(n);

/**
 * Strip leading quote marks ('INV-1), collapse stray spaces, expand Excel scientific
 * notation. Internal dashes and slashes are preserved ("INV/2026/01" stays as is).
 */
const normalizeInvoiceNumber = (raw) => {
  if (typeof raw === 'number') return Number.isFinite(raw) ? expandNumber(raw) : '';
  const text = String(raw ?? '')
    .replace(/[\u00a0\u200b\ufeff]/g, ' ')
    .trim()
    .replace(/^['‘’`´]+/, '')
    .trim();
  return SCIENTIFIC.test(text) ? expandNumber(Number(text)) : text.replace(/\s+/g, ' ');
};

/* ── GSTIN ───────────────────────────────────────────────────────── */
const normalizeGSTIN = (raw) => String(raw ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');

/* ── Money ───────────────────────────────────────────────────────── */
const BLANK_TOKENS = new Set(['', '-', '–', '—', 'nil', 'na', 'n/a', 'null']);

/**
 * Parse an amount from anything an accounting package can emit:
 *   "₹ 1,25,000.00"  "(500)"  "1,234.50-" (SAP)  "11,800.00 Dr" (Tally)  "1.5E+3"
 * Returns { value, status } where status is 'ok' | 'blank' | 'invalid'.
 */
const parseAmount = (raw) => {
  if (typeof raw === 'number') {
    return Number.isFinite(raw) ? { value: round2(raw), status: 'ok' } : { value: null, status: 'invalid' };
  }
  const text = String(raw ?? '').replace(/[\u00a0\u200b]/g, ' ').trim();
  if (BLANK_TOKENS.has(text.toLowerCase())) return { value: null, status: 'blank' };

  const noDrCr = text.replace(/(?<=[\d.)\s])\s*(dr|cr)\.?$/i, '').trim();        // Tally Dr/Cr suffix
  const negative = /^\(.*\)$/.test(noDrCr) || /^-/.test(noDrCr) || /-$/.test(noDrCr);
  const digits = noDrCr
    .replace(/^\(|\)$/g, '')
    .replace(/^-|-$/g, '')
    .replace(/(rs\.?|inr|usd)|[₹$€£,\s]/gi, '');

  if (digits === '') return { value: null, status: 'blank' };            // e.g. a lone "₹"
  if (!/^(\d*\.?\d+|\d+\.)(e[+-]?\d+)?$/i.test(digits)) return { value: null, status: 'invalid' };
  const value = Number(digits);
  return Number.isFinite(value)
    ? { value: round2(negative ? -value : value), status: 'ok' }
    : { value: null, status: 'invalid' };
};

/* ── Dates → YYYY-MM-DD ──────────────────────────────────────────── */
const MONTH_NUMBER = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12,
};
const pad2 = (n) => String(n).padStart(2, '0');
const daysInMonth = (y, m) => new Date(y, m, 0).getDate();
const toIso = (y, m, d) =>
  m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m) ? `${y}-${pad2(m)}-${pad2(d)}` : '';
const fullYear = (yy) => (yy >= 100 ? yy : yy < 50 ? 2000 + yy : 1900 + yy);
const monthFromName = (name) => MONTH_NUMBER[String(name).toLowerCase().slice(0, 4)] || MONTH_NUMBER[String(name).toLowerCase().slice(0, 3)] || 0;
const serialToIso = (n) => {
  const d = XLSX.SSF.parse_date_code(n);
  return d ? toIso(d.y, d.m, d.d) : '';
};

/**
 * Excel serial (45201), 02-10-2026, 2026/10/02, 2-Oct-26, Oct 2, 2026 … → "2026-10-02".
 * Indian files are day-first, so 02/10/2026 is 2 October. Returns '' when unreadable.
 */
const normalizeDate = (raw) => {
  if (raw instanceof Date) return toIso(raw.getFullYear(), raw.getMonth() + 1, raw.getDate());
  if (typeof raw === 'number') return raw >= 20000 && raw <= 80000 ? serialToIso(raw) : '';
  const text = String(raw ?? '').trim().replace(/^['‘’]/, '');
  if (!text) return '';

  const iso = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T\s].*)?$/.exec(text);
  if (iso) return toIso(+iso[1], +iso[2], +iso[3]);

  const dmy = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?:[T\s].*)?$/.exec(text);
  if (dmy) {
    const y = fullYear(+dmy[3]);
    return toIso(y, +dmy[2], +dmy[1]) || toIso(y, +dmy[1], +dmy[2]); // 10/25/2026 → month-first fallback
  }

  const dMonY = /^(\d{1,2})(?:st|nd|rd|th)?[-/.\s]+([A-Za-z]{3,9})\.?[-/.,\s]+(\d{2,4})$/.exec(text);
  if (dMonY && monthFromName(dMonY[2])) return toIso(fullYear(+dMonY[3]), monthFromName(dMonY[2]), +dMonY[1]);

  const monDY = /^([A-Za-z]{3,9})\.?[-/.\s]+(\d{1,2})(?:st|nd|rd|th)?[-/.,\s]+(\d{2,4})$/.exec(text);
  if (monDY && monthFromName(monDY[1])) return toIso(fullYear(+monDY[3]), monthFromName(monDY[1]), +monDY[2]);

  return /^\d{5}(\.\d+)?$/.test(text) ? serialToIso(Number(text)) : '';
};

/* ═══════════════ 2. Column alias dictionary + fuzzy mapper ═══════════════ */
/**
 * strong  = unambiguous names           (exact 1.00)
 * weak    = plausible but ambiguous      (exact 0.75, e.g. "Voucher No" is an internal number)
 * generic = one-word headers             (exact 0.55, e.g. "Amount" could be gross or taxable)
 * `reject` stops rate / non-taxable columns from being mistaken for amounts.
 */
const RATE_OR_PERCENT = /%|\brate\b|\bperc/i;
const FIELD_DEFS = {
  gstin: {
    label: 'GSTIN', required: true,
    strong: ['gstin', 'gstin/uin', 'party gstin', 'vendor gstin', 'gst no', 'gst number', 'supplier gstin', 'gstin of supplier', 'supplier gst no', 'gstin no', 'party gst no', 'party gst number', 'vendor gst no', 'vendor gst number'],
    weak: ['tin', 'uin'],
  },
  invoice: {
    label: 'Invoice Number', required: true,
    strong: ['invoice number', 'invoice no', 'invoice num', 'inv no', 'inv number', 'invoice #', 'bill no', 'bill number', 'bill #', 'supplier invoice no', 'supplier invoice number', 'supplier bill no', 'party invoice no', 'vendor invoice no', 'vendor bill no', 'invoice'],
    weak: ['vch no', 'vch number', 'voucher no', 'voucher number', 'doc no', 'doc number', 'document no', 'document number', 'ref no'],
    generic: ['bill'],
  },
  taxable: {
    label: 'Taxable Value', required: true,
    strong: ['taxable value', 'taxable amount', 'taxable amt', 'taxable val', 'assessable value', 'taxable', 'basic value', 'basic amount'],
    weak: ['subtotal', 'sub total'],
    generic: ['amount', 'value'],
    reject: /non[\s-]?taxable|exempt|nil[\s-]?rated|zero[\s-]?rated|%|\brate\b/i,
  },
  date: {
    label: 'Invoice Date', required: false,
    strong: ['invoice date', 'bill date', 'inv date', 'document date', 'supplier invoice date'],
    weak: ['vch date', 'voucher date', 'posting date'],
    generic: ['date'],
  },
  name: {
    label: 'Supplier Name', required: false,
    strong: ['party name', 'supplier name', 'vendor name', 'trade name', 'legal name', 'trade/legal name', 'name of supplier', 'name of party', 'party', 'supplier', 'vendor'],
    weak: ['particulars'],
    generic: ['name'],
  },
  igst: {
    label: 'IGST', required: false,
    strong: ['igst', 'integrated tax', 'igst amount', 'igst (₹)', 'igst amt', 'integrated tax amount'],
    reject: RATE_OR_PERCENT,
  },
  cgst: {
    label: 'CGST', required: false,
    strong: ['cgst', 'central tax', 'cgst amount', 'cgst (₹)', 'cgst amt', 'central tax amount'],
    reject: RATE_OR_PERCENT,
  },
  sgst: {
    label: 'SGST / UTGST', required: false,
    strong: ['sgst', 'state tax', 'state/ut tax', 'sgst amount', 'sgst (₹)', 'sgst amt', 'sgst/utgst', 'utgst', 'state tax amount'],
    reject: RATE_OR_PERCENT,
  },
  totalTax: {
    label: 'Total Tax', required: false,
    strong: ['total tax', 'tax amount', 'total tax (₹)', 'gst amount', 'total gst', 'total tax amount', 'tax amt'],
    reject: /taxable|%|\brate\b|\bperc/i,
  },
};

const MANDATORY = Object.keys(FIELD_DEFS).filter((f) => FIELD_DEFS[f].required);

const CONF = {
  exact: { strong: 1, weak: 0.75, generic: 0.55 },
  contains: { strong: 0.85, weak: 0.6 },
  fuzzy: { strong: 0.8, weak: 0.55 },
};

/** Short one-word aliases would match inside unrelated headers, so containment is opt-in. */
const NO_CONTAIN = new Set(['invoice', 'taxable', 'party', 'supplier', 'vendor', 'particulars', 'subtotal', 'bill']);
const ALWAYS_CONTAIN = new Set(['gstin', 'igst', 'cgst', 'sgst']);
const mayContain = (alias) => ALWAYS_CONTAIN.has(alias) || (alias.length >= 6 && !NO_CONTAIN.has(alias));

/** Flattened, normalised alias list for one field: [{ alias, tier, order }]. */
const aliasTable = (def) =>
  ['strong', 'weak', 'generic']
    .flatMap((tier) => (def[tier] || []).map((a) => ({ alias: normalizeHeader(a), tier })))
    .map((entry, order) => ({ ...entry, order }));

const ALIASES = Object.fromEntries(Object.entries(FIELD_DEFS).map(([f, def]) => [f, aliasTable(def)]));

/** Best confidence with which header text `n` could be field `field` (or null). */
const scoreAlias = (n, field) =>
  ALIASES[field]
    .map(({ alias, tier, order }) => {
      const conf =
        n === alias ? CONF.exact[tier]
        : tier !== 'generic' && mayContain(alias) && n.includes(alias) ? CONF.contains[tier]
        : tier !== 'generic' && n.length >= 5 && alias.length >= 5 && similarity(n, alias) >= 0.85 ? CONF.fuzzy[tier]
        : 0;
      return { conf, alias, order };
    })
    .filter((m) => m.conf > 0)
    .reduce((top, m) => (!top || m.conf > top.conf || (m.conf === top.conf && m.order < top.order) ? m : top), null);

/**
 * Map header cells to canonical fields. Greedy by confidence so each column and each
 * field is used once: { gstin: { index, header, confidence, alias }, … }
 */
const mapColumns = (headers) => {
  const candidates = headers.flatMap((h, index) => {
    const text = String(h ?? '').trim();
    const n = normalizeHeader(text);
    if (!n) return [];
    return Object.keys(FIELD_DEFS).flatMap((field) => {
      const reject = FIELD_DEFS[field].reject;
      if (reject && reject.test(text)) return [];
      const hit = scoreAlias(n, field);
      return hit && hit.conf >= MIN_USABLE ? [{ field, index, header: text, confidence: hit.conf, alias: hit.alias, order: hit.order }] : [];
    });
  });

  return candidates
    .sort((a, b) => b.confidence - a.confidence || a.order - b.order || a.index - b.index)
    .reduce(
      (acc, c) => {
        if (acc.fields.has(c.field) || acc.cols.has(c.index)) return acc;
        acc.fields.add(c.field);
        acc.cols.add(c.index);
        acc.map[c.field] = { index: c.index, header: c.header, confidence: c.confidence, alias: c.alias };
        return acc;
      },
      { fields: new Set(), cols: new Set(), map: {} }
    ).map;
};

/** Mandatory fields that are unmapped (missing) or mapped with low confidence. */
const weakMandatory = (map) =>
  MANDATORY.filter((f) => !map[f] || map[f].confidence < CONFIDENT);

/* ═══════════════════════ 3. Header-row detection ═══════════════════════ */
const KEYWORDS = ['gstin', 'invoice', 'vch', 'voucher', 'taxable', 'bill', 'party', 'supplier', 'vendor', 'amount', 'igst', 'cgst', 'sgst'];

/** Overlay two consecutive header rows (official GSTR-2B has a 2-row merged header). */
const mergeHeaderRows = (upper = [], lower = []) =>
  Array.from({ length: Math.max(upper.length, lower.length) }, (_, k) =>
    isFilled(lower[k]) ? lower[k] : upper[k]
  );

const scoreHeaderCandidate = (headers) => {
  const map = mapColumns(headers);
  const mapped = Object.values(map).reduce((t, m) => t + m.confidence, 0);
  const keywordHits = headers.filter(
    (h) => typeof h === 'string' && KEYWORDS.some((k) => normalizeHeader(h).includes(k))
  ).length;
  return { map, score: mapped + Math.min(keywordHits, 5) * 0.1 };
};

/**
 * Scan the first 20 rows. Each row is tried alone and merged with the row below it.
 * `startIndex` = first header row (rows above it are banners), `index` = last header row.
 */
const detectHeader = (rows) =>
  rows
    .slice(0, HEADER_SCAN_ROWS)
    .flatMap((row, i) => [
      { startIndex: i, index: i, headers: row },
      { startIndex: i, index: i + 1, headers: mergeHeaderRows(row, rows[i + 1]) },
    ])
    .filter((c) => !isBlankRow(c.headers))
    .filter((c) => c.index === c.startIndex || !isBlankRow(rows[c.startIndex]))   // merged pair needs a real upper row
    .reduce(
      (best, cand) => {
        const { map, score } = scoreHeaderCandidate(cand.headers);
        return score > best.score ? { ...cand, map, score } : best;
      },
      { startIndex: 0, index: 0, headers: rows[0] || [], map: {}, score: -1 }
    );

/* ═══════════════════ 4. Junk / footer row filtering ═══════════════════ */
const SUMMARY_CELL = /^(grand\s*total|sub\s*-?\s*total|totals?|carried\s*(over|forward)|brought\s*forward|c\/f|b\/f|closing\s*balance|opening\s*balance)\s*[:.\-]*$/i;
const PAGE_CELL = /^(page\s*(no\.?)?\s*\d+(\s*(of|\/)\s*\d+)?|\d+\s*of\s*\d+)$/i;
const STARTS_WITH_TOTAL = /^\s*(grand\s*)?(sub\s*-?\s*)?total\b/i;

const textCells = (row) => row.filter((c) => typeof c === 'string').map((c) => c.trim());

/** Whole-cell 'TOTAL' / 'GRAND TOTAL' / 'CARRIED OVER' / 'Page 2 of 3' — or "Total …" on a row with no GSTIN. */
const isSummaryRow = (row, gstin) =>
  textCells(row).some((c) => SUMMARY_CELL.test(c) || PAGE_CELL.test(c)) ||
  (!gstin && textCells(row).some((c) => STARTS_WITH_TOTAL.test(c)));

/** A header line printed again after a page break. */
const isRepeatedHeader = (row, headers) => {
  const known = headers.map(normalizeHeader);
  const nonEmpty = known.filter(Boolean).length;
  const same = known.filter((h, i) => h && normalizeHeader(row[i]) === h).length;
  return same >= Math.max(2, Math.ceil(nonEmpty * 0.5));
};

/* ═══════════════════ 5. Row normalisation ═══════════════════ */
const AMOUNT_FIELDS = ['taxable', 'igst', 'cgst', 'sgst', 'totalTax'];

/**
 * Value stored for reconcile.js validation:
 *   valid number → number, blank → '', unparseable text → the original text (so it's flagged)
 */
const amountField = (raw) => {
  const { value, status } = parseAmount(raw);
  return status === 'ok' ? value : status === 'blank' ? '' : String(raw).trim();
};

const normalizeRow = (map) => (row) => {
  const cell = (field) => (map[field] ? row[map[field].index] : undefined);
  const date = normalizeDate(cell('date'));
  return {
    gstin: normalizeGSTIN(cell('gstin')),
    invoice: normalizeInvoiceNumber(cell('invoice')),
    date: date || String(cell('date') ?? '').trim(), // keep unreadable dates visible rather than losing them
    name: String(cell('name') ?? '').trim(),
    ...Object.fromEntries(AMOUNT_FIELDS.map((f) => [f, amountField(cell(f))])),
  };
};

/** Classify each data row once, so the counts reported to the user are exact. */
const classifyRows = (dataRows, headers, map) =>
  dataRows.map(({ raw, rowNumber }) => {
    if (isBlankRow(raw)) return { kind: 'empty', rowNumber };
    if (isRepeatedHeader(raw, headers)) return { kind: 'repeatedHeader', rowNumber };
    const fields = normalizeRow(map)(raw);
    if (isSummaryRow(raw, fields.gstin)) return { kind: 'summary', rowNumber };
    if (!fields.gstin && !fields.invoice) return { kind: 'blankKeys', rowNumber };
    return { kind: 'keep', rowNumber, raw, fields };
  });

/* ═══════════════════ 6. Manual-mapping support ═══════════════════ */
const FIELD_LIST = Object.entries(FIELD_DEFS).map(([key, d]) => ({ key, label: d.label, required: d.required }));

const displayHeaders = (headers, width) =>
  Array.from({ length: width }, (_, i) => String(headers[i] ?? '').trim() || `Column ${colLetter(i)}`);

const previewOf = (rows) =>
  rows.slice(0, PREVIEW_ROWS).map((row) =>
    Array.from({ length: Math.min(PREVIEW_COLS, row.length) }, (_, i) => String(row[i] ?? '').slice(0, 60))
  );

/** The JSON the UI needs to render the "map your columns" modal. */
const buildMappingRequest = ({ side, fileName, sheet, rows, header }) => {
  const preview = previewOf(rows);
  const width = preview.reduce((w, r) => Math.max(w, r.length), 0);
  const weak = weakMandatory(header.map);
  return {
    side,
    label: SIDE_LABEL[side],
    fileName,
    sheet,
    headerRowIndex: header.score >= 1 ? header.index : Math.max(0, rows.findIndex((r) => textCells(r).length >= 3)),
    detectedHeaders: displayHeaders(header.headers, width),
    rawPreviewRows: preview,
    mapping: Object.fromEntries(Object.entries(header.map).map(([f, m]) => [f, { index: m.index, confidence: m.confidence }])),
    missingFields: MANDATORY.filter((f) => !header.map[f]),
    lowConfidenceFields: MANDATORY.filter((f) => header.map[f] && header.map[f].confidence < CONFIDENT),
    fields: FIELD_LIST,
  };
};

/**
 * Turn the manual-mapping form (`books_headerRow`, `books_col_gstin`, …) into overrides.
 * Returns { present:false } when that file wasn't part of the form.
 */
const buildOverrides = (form, side) => {
  if (form[`${side}_headerRow`] === undefined) return { present: false };
  const headerRowIndex = Number(form[`${side}_headerRow`]);
  const columns = Object.fromEntries(
    Object.keys(FIELD_DEFS).map((f) => {
      const v = form[`${side}_col_${f}`];
      return [f, v === undefined || v === '' ? null : Number(v)];
    })
  );
  const chosen = Object.values(columns).filter((v) => v !== null);
  const requiredCols = MANDATORY.map((f) => columns[f]);
  const error =
    !Number.isInteger(headerRowIndex) || headerRowIndex < 0 ? 'Please pick the header row.'
    : chosen.some((v) => !Number.isInteger(v) || v < 0) ? 'One of the chosen columns is not valid.'
    : MANDATORY.find((f) => columns[f] === null) ? `Please choose the ${FIELD_DEFS[MANDATORY.find((f) => columns[f] === null)].label} column for ${SIDE_LABEL[side]}.`
    : new Set(requiredCols).size !== requiredCols.length ? `${SIDE_LABEL[side]}: GSTIN, Invoice Number and Taxable Value need three different columns.`
    : null;
  return { present: true, headerRowIndex, columns, error };
};

/** Auto-detected mapping with the user's explicit choices laid on top (null = "not in my file"). */
const applyOverrides = (autoMap, headers, columns) => {
  const manual = Object.entries(columns)
    .filter(([, i]) => Number.isInteger(i))
    .map(([field, index]) => [field, { index, header: String(headers[index] ?? '').trim(), confidence: 1, alias: 'manual' }]);
  const manualFields = new Set(Object.keys(columns));
  const manualCols = new Set(manual.map(([, m]) => m.index));
  const kept = Object.entries(autoMap).filter(([f, m]) => !manualFields.has(f) && !manualCols.has(m.index));
  return Object.fromEntries([...kept, ...manual]);
};

/* ═══════════════════════ 7. The entry point ═══════════════════════ */
const readRows = (workbook, name) =>
  // range:0 + blankrows:true → array index + 1 is the real Excel row number
  XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', blankrows: true, range: 0 });

/**
 * Normalise one uploaded file.
 *   { ok:false, error }                      unreadable / empty file
 *   { ok:false, needsManualMapping, request } mandatory columns not found with confidence
 *   { ok:true, rows, headers, mapping, stats, notes, … }
 * `overrides` = { headerRowIndex, columns } from the manual-mapping modal.
 */
const normalizeExcelData = (file, { side = 'books', overrides } = {}) => {
  const label = SIDE_LABEL[side];
  const fileName = file.originalname;

  let workbook;
  try {
    workbook = XLSX.read(file.buffer, { type: 'buffer', raw: true });
  } catch (e) {
    return { ok: false, error: `Couldn't read the ${label} file "${fileName}". Make sure it is a valid .xlsx, .xls or .csv file.` };
  }

  const sheets = workbook.SheetNames
    .map((name) => {
      const rows = readRows(workbook, name);
      return { name, rows, header: detectHeader(rows), filled: rows.filter((r) => !isBlankRow(r)).length };
    })
    .sort((a, b) =>
      b.header.score - a.header.score ||
      Number(/^b2b$/i.test(b.name)) - Number(/^b2b$/i.test(a.name)) ||
      b.filled - a.filled
    );

  const best = sheets[0];
  if (!best || best.filled === 0) {
    return { ok: false, error: `The ${label} file "${fileName}" looks empty. Please upload a file that has data rows.` };
  }

  const { rows } = best;
  const manual = Boolean(overrides && Number.isInteger(overrides.headerRowIndex));
  const headerIndex = manual ? overrides.headerRowIndex : best.header.index;
  const startIndex = manual ? overrides.headerRowIndex : best.header.startIndex;
  const headers = manual ? rows[headerIndex] || [] : best.header.headers;
  const mapping = manual
    ? applyOverrides(mapColumns(headers), headers, overrides.columns)
    : best.header.map;

  if (!manual && weakMandatory(mapping).length > 0) {
    return {
      ok: false,
      needsManualMapping: true,
      request: buildMappingRequest({ side, fileName, sheet: best.name, rows, header: best.header }),
    };
  }

  const classified = classifyRows(
    rows.slice(headerIndex + 1).map((raw, i) => ({ raw, rowNumber: headerIndex + i + 2 })),
    headers,
    mapping
  );
  const count = (kind) => classified.filter((c) => c.kind === kind).length;
  const kept = classified.filter((c) => c.kind === 'keep');

  const removedSummaryRows = count('summary');
  const removedBlankRows = count('blankKeys');
  const repeatedHeaderRows = count('repeatedHeader');
  const stats = {
    topRowsSkipped: startIndex,
    columnsNormalized: Object.keys(mapping).length,
    removedRows: removedSummaryRows + removedBlankRows + repeatedHeaderRows,
    removedSummaryRows,
    removedBlankRows,
    repeatedHeaderRows,
    rowsKept: kept.length,
  };

  // Tell the user when a mandatory column was matched on a weak guess (e.g. Voucher No → Invoice).
  const notes = MANDATORY
    .filter((f) => mapping[f] && mapping[f].confidence < 1 && mapping[f].alias !== 'manual')
    .map((f) => `${FIELD_DEFS[f].label} was read from the column "${mapping[f].header}". Check it is the right one.`);

  return {
    ok: true,
    side,
    label,
    fileName,
    sheet: best.name,
    manual,
    headers,
    headerIndex,
    mapping,
    rows: kept.map(({ raw, rowNumber, fields }) => ({ rowNumber, raw, fields })),
    stats,
    notes,
  };
};

/** "Auto-cleaned [File]: Skipped X top header rows, normalized Y column names, removed Z summary rows." */
const cleaningSummary = ({ fileName, stats }) =>
  `Auto-cleaned ${fileName}: Skipped ${plural(stats.topRowsSkipped, 'top header row')}, ` +
  `normalized ${plural(stats.columnsNormalized, 'column name')}, ` +
  `removed ${plural(stats.removedRows, 'summary row')}.`;

module.exports = {
  // pipeline
  normalizeExcelData, cleaningSummary,
  // header + mapping
  detectHeader, mapColumns, scoreAlias, weakMandatory, FIELD_DEFS, FIELD_LIST, MANDATORY,
  // manual mapping
  buildMappingRequest, buildOverrides, applyOverrides,
  // cell scrubbing
  normalizeInvoiceNumber, normalizeGSTIN, parseAmount, normalizeDate,
  // row filtering
  isSummaryRow, isRepeatedHeader,
  // shared
  normalizeHeader, round2, isFilled, isBlankRow, CONFIDENT,
};
