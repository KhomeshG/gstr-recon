'use strict';
const XLSX = require('xlsx-js-style'); // drop-in SheetJS fork that supports cell colours

const COLORS = {
  exact:    { head: '2E7D32', row: 'E3F4E5' },
  mismatch: { head: 'B7791F', row: 'FFF6D6' },
  portal:   { head: 'C62828', row: 'FDE3E3' },
  books:    { head: '1565C0', row: 'E1EEFC' },
  summary:  { head: '14213D', row: 'F1F3F7' },
  late:     { head: '6D28D9', row: 'EEE7FB' },
  skipped:  { head: '475467', row: 'F2F4F7' },
};

const side = { style: 'thin', color: { rgb: 'D0D5DD' } };
const borders = { top: side, bottom: side, left: side, right: side };
const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;

const headerCell = (v, color) => ({
  v, t: 's',
  s: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: color } }, border: borders, alignment: { vertical: 'center', wrapText: true } },
});

const bodyCell = (v, color) => {
  const isNum = typeof v === 'number';
  return {
    v: v ?? '',
    t: isNum ? 'n' : 's',
    s: { fill: { fgColor: { rgb: color } }, border: borders, ...(isNum ? { numFmt: '#,##0.00' } : {}) },
  };
};

const buildSheet = (headers, rows, palette) => {
  const aoa = [
    headers.map((h) => headerCell(h, palette.head)),
    ...rows.map((r) => r.map((v) => bodyCell(v, palette.row))),
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = headers.map((h, i) => ({
    wch: Math.min(42, Math.max(h.length + 2, ...rows.map((r) => String(r[i] ?? '').length + 2))),
  }));
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(rows.length, 1), c: headers.length - 1 } }) };
  return ws;
};

const PLAIN_HEAD = ['Supplier', 'GSTIN', 'Invoice No', 'Date', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Total Tax'];
const plainRow = (r) => [r.name, r.gstin, r.invoiceRaw, r.date, r.taxable, r.igst, r.cgst, r.sgst, r.tax];

const summarySheet = (s) =>
  buildSheet(
    ['Metric', 'Count', 'Tax Amount (₹)'],
    [
      ['Purchase Register invoices', s.booksCount, ''],
      ['GSTR-2B invoices', s.portalCount, ''],
      ['Exact matches', s.exact.count, s.exact.tax],
      ['Mismatches / probable typos', s.mismatches.count, s.mismatches.tax],
      ['Missing in GSTR-2B (unclaimed ITC)', s.missingInPortal.count, s.missingInPortal.tax],
      ['Missing in Books', s.missingInBooks.count, s.missingInBooks.tax],
      ['Late claimed ITC (prior-month match)', s.lateClaimed?.count ?? 0, s.lateClaimed?.tax ?? 0],
      ['Skipped / invalid rows', s.skippedCount ?? 0, ''],
      ['ITC at risk', '', s.itcAtRisk],
      ['Match rate', `${s.matchRate}%`, ''],
    ],
    COLORS.summary
  );

/** Returns an .xlsx Buffer with Summary + 4 colour-coded sheets. */
const NON_COMPLIANT_HEAD = ['Supplier Name', 'GSTIN', 'Invoice No', 'Date', 'Taxable Value', 'Total Tax (ITC at Risk)'];

/** Follow-up sheet: every invoice that is in the books but missing from GSTR-2B. */
const buildNonCompliantBuffer = (missingInPortal) => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildSheet(
    NON_COMPLIANT_HEAD,
    missingInPortal.map((r) => [r.name, r.gstin, r.invoiceRaw, r.date, r.taxable, r.tax]),
    COLORS.portal
  ), 'Non-Compliant Suppliers');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
};

const rawText = (raw) => Object.entries(raw).map(([k, v]) => `${k}: ${v}`).join(' | ');

const buildExportBuffer = ({ results, summary, skippedRows = [] }) => {
  const wb = XLSX.utils.book_new();
  const add = (name, ws) => XLSX.utils.book_append_sheet(wb, ws, name);

  add('Summary', summarySheet(summary));
  add('Exact Matches', buildSheet(PLAIN_HEAD, results.exact.map((m) => plainRow(m.books)), COLORS.exact));
  add('Mismatches', buildSheet(
    ['GSTIN', 'Books Invoice', 'Portal Invoice', 'Books Taxable', 'Portal Taxable', 'Taxable Diff', 'Books Tax', 'Portal Tax', 'Tax Diff', 'Reason'],
    results.mismatches.map(({ books: b, portal: p, reason }) => [
      b.gstin, b.invoiceRaw, p.invoiceRaw, b.taxable, p.taxable, round2(b.taxable - p.taxable),
      b.tax, p.tax, round2(b.tax - p.tax), reason,
    ]),
    COLORS.mismatch
  ));
  add('Missing in GSTR-2B', buildSheet(PLAIN_HEAD, results.missingInPortal.map(plainRow), COLORS.portal));
  add('Missing in Books', buildSheet(PLAIN_HEAD, results.missingInBooks.map(plainRow), COLORS.books));

  add('Late Claimed ITC', buildSheet(
    ['Supplier', 'GSTIN', 'Invoice No', 'Original Period', 'Date', 'Taxable Value', 'Total Tax', 'Note'],
    (results.lateClaimed || []).map(({ ledger: l, portal: p, message }) => [p.name || l.name, p.gstin, p.invoiceRaw, l.periodLabel, p.date, p.taxable, p.tax, message]),
    COLORS.late
  ));
  add('Skipped Rows', buildSheet(
    ['Source', 'Excel Row', 'Reason', 'Raw Data'],
    skippedRows.map((s) => [s.sourceLabel, s.rowIndex, s.reasonForSkipping, rawText(s.rawData)]),
    COLORS.skipped
  ));

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
};

module.exports = { buildExportBuffer, buildNonCompliantBuffer };
