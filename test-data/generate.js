'use strict';
/**
 * Generates a hard GSTR-2B reconciliation test pack:
 *  1. GSTR2B_Official_Format_Sep2026.xlsx       (portal layout: Read me + B2B with 2-row merged header)
 *  2. Purchase_Register_Tally_Style_Sep2026.xlsx (preamble rows, messy cells, junk rows, total row)
 *  3. Answer_Key_Expected_Results.xlsx           (what each invoice SHOULD land in)
 * Run: node test-data/generate.js
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx-js-style');

const OUT = path.join(__dirname, 'files');
fs.mkdirSync(OUT, { recursive: true });

const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const SUP = {
  s1: ['Shree Traders', '27AABCS1234A1Z5'],
  s2: ['Nagpur Steel Co', '27AAACN5678B1Z2'],
  s3: ['Orange City Papers', '27AAECO9012C1Z8'],
  s4: ['Vidarbha Hardware', '27AAFCV3456D1Z1'],
  s5: ['Metro Packaging', '24AAGCM7890E1Z4'],
  s6: ['Kisan Agro Supplies', '27AAHCK1122F1Z9'],
  s7: ['Sai Logistics', '27AAICS3344G1Z3'],
  s8: ['Bharat Electricals', '29AAJCB5566H1Z7'],
  s9: ['Capital Stationers', '07AAKCC7788J1Z6'],
  s10: ['Ganga Chemicals', '24AALCG9900K1Z0'],
};
const INTER = new Set(['s5', 's8', 's9', 's10']);

const taxOf = (taxable, rate, inter) => {
  const t = r2((taxable * rate) / 100);
  return inter ? { igst: t, cgst: 0, sgst: 0 } : { igst: 0, cgst: r2(t / 2), sgst: r2(t / 2) };
};

/* side = one invoice row as it appears in a file */
const side = (sup, base, ov = {}) => {
  const m = { ...base, ...ov };
  const [name, gstin] = SUP[sup];
  const tax = m.tax || taxOf(m.taxable, m.rate ?? 18, INTER.has(sup));
  return { name, gstin: m.gstin ?? gstin, inv: m.inv, taxable: m.taxable, taxableCell: m.taxableCell, tax, rate: m.rate ?? 18, sup };
};

const entries = [];
const pair = (id, exp, tests, sup, base, bOv, pOv) =>
  entries.push({ id, exp, tests, books: side(sup, base, bOv), portal: side(sup, base, pOv) });
const booksOnly = (id, exp, tests, sup, base, bOv) =>
  entries.push({ id, exp, tests, books: side(sup, base, bOv) });
const portalOnly = (id, exp, tests, sup, base, pOv) =>
  entries.push({ id, exp, tests, portal: side(sup, base, pOv) });

/* ─────────────── The test cases ─────────────── */
pair('T01', 'EXACT', 'Plain identical invoice', 's1', { inv: 'INV/2026/001', taxable: 10000 });
pair('T02', 'EXACT', 'Invoice no. punctuation/case differs (inv-2026-002 vs INV/2026/002)', 's1', { inv: 'INV/2026/002', taxable: 25000 }, { inv: 'inv-2026-002' });
pair('T03', 'EXACT', 'GSTIN lower-case with stray spaces + invoice spacing', 's2', { inv: 'NS-1042', taxable: 50000 }, { gstin: ' 27 aaacn5678b1z2 ', inv: 'ns 1042' });
pair('T04', 'EXACT', 'Invoice no. stored as a NUMBER in books, TEXT on portal', 's3', { inv: '1045', taxable: 7500.5 }, { inv: 1045 });
pair('T05', 'EXACT', 'Books amount typed as text "₹ 1,25,000.00" (Indian comma style)', 's4', { inv: 'VH/2026/0210', taxable: 125000 }, { taxableCell: '₹ 1,25,000.00' });
pair('T06', 'EXACT', 'Floating-point noise 15000.004 vs 15000.00 (rounds to 2 dp), IGST', 's5', { inv: 'MP-5521', taxable: 15000 }, { taxable: 15000.004 });
pair('T07a', 'EXACT', 'Same supplier, identical amounts, different invoices (no cross-pairing)', 's6', { inv: 'KA/331', taxable: 4200 });
pair('T07b', 'EXACT', 'Same supplier, identical amounts, different invoices (no cross-pairing)', 's6', { inv: 'KA/332', taxable: 4200 });
pair('T08', 'EXACT', 'Leading zeros kept in invoice no. (SL/00908)', 's7', { inv: 'SL/00908', taxable: 8800 });
pair('T09', 'EXACT', '5% GST rate, small paise value, IGST', 's8', { inv: 'BE-77', taxable: 880.75, rate: 5 });
pair('T10a', 'EXACT', 'DUPLICATE in books: first copy should match the single portal invoice', 's9', { inv: 'CS/501', taxable: 12000 });
booksOnly('T10b', 'MISSING_PORTAL', 'DUPLICATE in books: second copy must NOT match again', 's9', { inv: 'CS/501', taxable: 12000 });
pair('T11', 'PROBABLE', 'Typo: letter O instead of zero (INV2026O12 vs INV2026012)', 's1', { inv: 'INV2026012', taxable: 18000 }, { inv: 'INV2026O12' });
pair('T12', 'PROBABLE', 'Leading zeros (0045 vs 45) and value differs by ₹3', 's3', { inv: '45', taxable: 9003 }, { inv: '0045', taxable: 9000 });
pair('T13', 'PROBABLE', 'Missing digit (BILL-2209 vs BILL-209)', 's2', { inv: 'BILL-209', taxable: 33000 }, { inv: 'BILL-2209' });
pair('T14', 'PROBABLE', 'Boundary: value differs by EXACTLY ₹10.00 and invoice differs', 's4', { inv: 'VH-2026-0301', taxable: 40010 }, { inv: 'VH/2026/0300', taxable: 40000 });
booksOnly('T15a', 'MISSING_PORTAL', 'Boundary: differs by ₹10.01 → must NOT be treated as probable', 's10', { inv: 'GC/88', taxable: 5000 });
portalOnly('T15b', 'MISSING_BOOKS', 'Boundary: differs by ₹10.01 → must NOT be treated as probable', 's10', { inv: 'GC/89', taxable: 5010.01 });
pair('T16a', 'PROBABLE', 'Two look-alike typos same supplier & amount: each pairs with the closest', 's6', { inv: 'KA/41O', taxable: 6000 }, { inv: 'KA/410' });
pair('T16b', 'PROBABLE', 'Two look-alike typos same supplier & amount: each pairs with the closest', 's6', { inv: 'KA/49O', taxable: 6000 }, { inv: 'KA/490' });
pair('T17', 'VALUE_MISMATCH', 'Same GSTIN + invoice, taxable differs by ₹1,000', 's2', { inv: 'NS/2088', taxable: 59000 }, { taxable: 60000 });
pair('T18', 'VALUE_MISMATCH', 'Same GSTIN + invoice, taxable differs by only ₹0.50', 's7', { inv: 'SL/1200', taxable: 10000 }, { taxable: 10000.5 });
pair('T19', 'EXACT*', 'Taxable matches but TAX differs by ₹180 (engine matches on taxable only)', 's8', { inv: 'BE-120', taxable: 20000 }, { tax: { igst: 3780, cgst: 0, sgst: 0 } }, { tax: { igst: 3600, cgst: 0, sgst: 0 } });
booksOnly('T20', 'MISSING_PORTAL', 'Supplier has not filed GSTR-1 (unclaimed ITC)', 's6', { inv: 'KA/777', taxable: 22000 });
booksOnly('T21a', 'MISSING_PORTAL', 'GSTIN typed wrong in books (last char Z1 instead of Z7)', 's8', { inv: 'BE-90', taxable: 14000 }, { gstin: '29AAJCB5566H1Z1' });
portalOnly('T21b', 'MISSING_BOOKS', 'Same invoice on portal under the correct GSTIN', 's8', { inv: 'BE-90', taxable: 14000 });
booksOnly('T22', 'MISSING_PORTAL', 'Negative amount (credit note booked as minus)', 's9', { inv: 'CN/12', taxable: -2500 });
portalOnly('T23', 'MISSING_BOOKS', 'Large invoice on portal, never booked', 's10', { inv: 'GC/120', taxable: 70000 });
portalOnly('T24', 'MISSING_BOOKS', 'Portal invoice never booked (note SL/908 vs booked SL/00908)', 's7', { inv: 'SL/908', taxable: 8800 });
portalOnly('T25', 'MISSING_BOOKS', 'Tiny ₹100 invoice on portal only', 's3', { inv: 'OCP/9001', taxable: 100 });

/* ─────────────── Cell helpers ─────────────── */
const num = (v) => ({ v, t: 'n', z: '#,##0.00' });
const str = (v) => ({ v: String(v), t: 's' });
const serial = (d, m, y) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
const bold = (v, fill = 'D9E1F2') => ({ v, t: 's', s: { font: { bold: true }, fill: { fgColor: { rgb: fill } } } });
const dayOf = (i) => (i % 28) + 1;
const widths = (n, w = 18) => Array.from({ length: n }, () => ({ wch: w }));

/* ─────────────── 1. Purchase register (Tally style) ─────────────── */
const bookEntries = entries.filter((e) => e.books).map((e, i) => ({ ...e, day: dayOf(i) }));
const booksAoa = [
  [bold('ABC Enterprises Pvt Ltd, Nagpur', 'FFFFFF')],
  [bold('Purchase Register', 'FFFFFF')],
  [str('1-Sep-2026 to 30-Sep-2026')],
  [],
  ['Date', 'Voucher No.', 'Supplier Invoice No.', 'Party Name', 'GSTIN/UIN', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Gross Total'].map((h) => bold(h)),
];
const bookRow = (e, n) => {
  const b = e.books;
  const taxableNum = typeof b.taxable === 'number' ? b.taxable : 0;
  return [
    str(`${String(e.day).padStart(2, '0')}-09-2026`),
    str(`PV/${String(n + 1).padStart(4, '0')}`),
    typeof b.inv === 'number' ? { v: b.inv, t: 'n' } : str(b.inv),
    str(b.name),
    str(b.gstin),
    b.taxableCell ? str(b.taxableCell) : num(b.taxable),
    num(b.tax.igst), num(b.tax.cgst), num(b.tax.sgst),
    num(r2(taxableNum + b.tax.igst + b.tax.cgst + b.tax.sgst)),
  ];
};
bookEntries.forEach((e, n) => {
  booksAoa.push(bookRow(e, n));
  if (n === 4) booksAoa.push([]); // blank row mid-data
  if (n === 9) booksAoa.push([str('05-09-2026'), str('PV/JUNK1'), str(''), str('Shree Traders'), str('27AABCS1234A1Z5'), num(500), num(0), num(45), num(45), num(590)]); // missing invoice no.
  if (n === 15) booksAoa.push([str('12-09-2026'), str('PV/JUNK2'), str('K-1'), str('Local Kirana (unregistered)'), str(''), num(1200), num(0), num(0), num(0), num(1200)]); // no GSTIN
});
booksAoa.push([str(''), str(''), str(''), bold('Grand Total', 'FFFFFF'), str(''), num(999999), num(1), num(1), num(1), num(1000002)]); // total row
const booksWs = XLSX.utils.aoa_to_sheet(booksAoa);
booksWs['!cols'] = widths(10);
const booksWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(booksWb, booksWs, 'Purchase Register');
XLSX.writeFile(booksWb, path.join(OUT, 'Purchase_Register_Tally_Style_Sep2026.xlsx'));

/* ─────────────── 2. Official-format GSTR-2B ─────────────── */
const portalEntries = entries
  .filter((e) => e.portal)
  .map((e, i) => ({ ...e, day: dayOf(i * 3) }))
  .sort((a, b) => a.portal.gstin.localeCompare(b.portal.gstin) || String(a.portal.inv).localeCompare(String(b.portal.inv)));

const readMe = [
  ['Goods and Services Tax - GSTR-2B'],
  [],
  ['GSTIN', '27ABCDE1234F1Z5'],
  ['Legal name of the registered person', 'ABC Enterprises Pvt Ltd'],
  ['Trade name, if any', 'ABC Enterprises'],
  ['Date of generation', '14/10/2026'],
  ['Return period', 'September 2026'],
  [],
  ['Read me: this is an auto-drafted ITC statement. Sheets: ITC Summary, B2B, B2B-CDNR.'],
].map((r) => r.map((c) => (c === '' ? c : str(c))));

const h1 = ['GSTIN of supplier', 'Trade/Legal name', 'Invoice details', '', '', '', '', 'Supply Attract Reverse Charge', 'Rate(%)', 'Taxable Value (₹)', 'Tax Amount', '', '', '', 'GSTR-1/IFF/GSTR-5 Period', 'GSTR-1/IFF/GSTR-5 Filing Date', 'ITC Availability'];
const h2 = ['', '', 'Invoice number', 'Invoice type', 'Invoice Date', 'Invoice Value(₹)', 'Place of supply', '', '', '', 'Integrated Tax(₹)', 'Central Tax(₹)', 'State/UT Tax(₹)', 'Cess(₹)', '', '', ''];
const b2b = [h1.map((c) => bold(c, 'BDD7EE')), h2.map((c) => bold(c, 'BDD7EE'))];
portalEntries.forEach((e, n) => {
  const p = e.portal;
  b2b.push([
    str(p.gstin), str(p.name),
    typeof p.inv === 'number' ? { v: p.inv, t: 'n' } : str(p.inv),
    str('Regular'),
    { v: serial(e.day, 9, 2026), t: 'n', z: 'dd-mm-yyyy' },
    num(r2(p.taxable + p.tax.igst + p.tax.cgst + p.tax.sgst)),
    str(p.gstin.startsWith('27') ? '27-Maharashtra' : '27-Maharashtra'),
    str('No'), { v: p.rate, t: 'n' }, num(p.taxable),
    num(p.tax.igst), num(p.tax.cgst), num(p.tax.sgst), num(0),
    str('Sep-2026'), str('11-10-2026'), str('Yes'),
  ]);
  if (n === 7) b2b.push([]); // stray blank row
});
const b2bWs = XLSX.utils.aoa_to_sheet(b2b);
b2bWs['!merges'] = [{ s: { r: 0, c: 2 }, e: { r: 0, c: 6 } }, { s: { r: 0, c: 10 }, e: { r: 0, c: 13 } }];
b2bWs['!cols'] = widths(17, 20);

const cdnr = XLSX.utils.aoa_to_sheet([
  ['GSTIN of supplier', 'Trade/Legal name', 'Note number', 'Note type', 'Note date', 'Taxable Value (₹)'].map((c) => bold(c)),
  [str('27AABCS1234A1Z5'), str('Shree Traders'), str('CN-1'), str('Credit note'), str('20-09-2026'), num(1000)],
]);
const summaryWs = XLSX.utils.aoa_to_sheet([[str('ITC Summary (auto-drafted)')], [str('Part A - ITC Available'), str('IGST'), str('CGST'), str('SGST')], [str('Other documents'), num(0), num(0), num(0)]]);

const portalWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(portalWb, XLSX.utils.aoa_to_sheet(readMe), 'Read me');
XLSX.utils.book_append_sheet(portalWb, summaryWs, 'ITC Summary');
XLSX.utils.book_append_sheet(portalWb, b2bWs, 'B2B');
XLSX.utils.book_append_sheet(portalWb, cdnr, 'B2B-CDNR');
XLSX.writeFile(portalWb, path.join(OUT, 'GSTR2B_Official_Format_Sep2026.xlsx'));

/* ─────────────── 3. Answer key ─────────────── */
const FILL = { EXACT: 'E3F4E5', 'EXACT*': 'E3F4E5', PROBABLE: 'FFF6D6', VALUE_MISMATCH: 'FFF6D6', MISSING_PORTAL: 'FDE3E3', MISSING_BOOKS: 'E1EEFC' };
const taxTotal = (s) => r2(s.tax.igst + s.tax.cgst + s.tax.sgst);
const keyRows = entries.map((e) => {
  const cell = (v, t) => ({ v: v ?? '', t: t || (typeof v === 'number' ? 'n' : 's'), s: { fill: { fgColor: { rgb: FILL[e.exp] } } } });
  return [cell(e.id), cell(e.exp), cell(e.tests), cell(e.books ? String(e.books.inv) : '—'), cell(e.portal ? String(e.portal.inv) : '—'), cell(e.books ? e.books.taxable : ''), cell(e.portal ? e.portal.taxable : '')];
});
const keyWs = XLSX.utils.aoa_to_sheet([
  ['Test ID', 'Expected bucket', 'What it tests', 'Invoice in books', 'Invoice on portal', 'Books taxable', 'Portal taxable'].map((h) => bold(h)),
  ...keyRows,
]);
keyWs['!cols'] = [{ wch: 8 }, { wch: 18 }, { wch: 78 }, { wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 14 }];

const count = (k) => entries.filter((e) => e.exp === k).length;
const expected = {
  booksUsable: entries.filter((e) => e.books).length,
  portalUsable: entries.filter((e) => e.portal).length,
  exact: count('EXACT') + count('EXACT*'),
  probable: count('PROBABLE'),
  value: count('VALUE_MISMATCH'),
  missingPortal: count('MISSING_PORTAL'),
  missingBooks: count('MISSING_BOOKS'),
  skippedBooksRows: 2,
  footerRowsRemoved: 1,
  itcAtRisk: r2(
    entries.filter((e) => e.exp === 'MISSING_PORTAL').reduce((t, e) => t + taxTotal(e.books), 0) +
    entries.filter((e) => ['PROBABLE', 'VALUE_MISMATCH'].includes(e.exp)).reduce((t, e) => t + Math.max(0, taxTotal(e.books) - taxTotal(e.portal)), 0)
  ),
};
const sumWs = XLSX.utils.aoa_to_sheet([
  ['Expected totals when both files are uploaded', 'Value'].map((h) => bold(h)),
  ['Purchase register invoices (usable rows)', expected.booksUsable],
  ['GSTR-2B invoices', expected.portalUsable],
  ['Exact matches (green)', expected.exact],
  ['Probable / typo matches (yellow)', expected.probable],
  ['Value mismatches (yellow)', expected.value],
  ['Mismatches tab total', expected.probable + expected.value],
  ['Missing in GSTR-2B (red)', expected.missingPortal],
  ['Missing in Books (blue)', expected.missingBooks],
  ['Skipped junk rows in books (warning banner)', expected.skippedBooksRows],
  ['Footer rows auto-removed ("Grand Total" line)', expected.footerRowsRemoved],
  ['ITC at risk (₹)', expected.itcAtRisk],
  [],
  ['* T19 is a known limitation: matching uses taxable value, so a tax-only difference stays green.'],
].map((r) => r.map((c) => (typeof c === 'string' ? str(c) : typeof c === 'number' ? { v: c, t: 'n' } : c))));
sumWs['!cols'] = [{ wch: 52 }, { wch: 14 }];
const keyWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(keyWb, sumWs, 'Expected totals');
XLSX.utils.book_append_sheet(keyWb, keyWs, 'Per-invoice key');
XLSX.writeFile(keyWb, path.join(OUT, 'Answer_Key_Expected_Results.xlsx'));

fs.writeFileSync(path.join(OUT, 'expected.json'), JSON.stringify({ expected, entries: entries.map((e) => ({ id: e.id, exp: e.exp, b: e.books && String(e.books.inv), p: e.portal && String(e.portal.inv) })) }, null, 2));
console.log('Generated into', OUT, expected);
