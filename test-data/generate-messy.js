'use strict';
/**
 * Messy-export test pack. The SAME six invoices are written in five different "dirty" shapes
 * (Tally, Zoho, SAP, hand-made client file, unmappable headers). After normalisation every one
 * must reconcile identically against the clean portal file.   Run: node test-data/generate-messy.js
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx-js-style');

const OUT = path.join(__dirname, 'messy');
fs.mkdirSync(OUT, { recursive: true });

// name, gstin, invoice, [d, m, y], taxable, igst, cgst, sgst
const TRUTH = [
  ['Shree Traders',       '27AABCS1234A1Z5', 'INV/2026/101', [3, 10, 2026],  10000,   0,     900,   900],
  ['Nagpur Steel Co',     '27AAACN5678B1Z2', 'NS-1042',      [5, 10, 2026],  50000,   0,     4500,  4500],
  ['Metro Packaging',     '24AAGCM7890E1Z4', '101000',       [8, 10, 2026],  15000,   2700,  0,     0],
  ['Kisan Agro Supplies', '27AAHCK1122F1Z9', 'KA/0331',      [12, 10, 2026], 4200,    0,     378,   378],
  ['Sai Logistics',       '27AAICS3344G1Z3', 'SL/908',       [15, 10, 2026], 125000,  0,     11250, 11250],
  ['Bharat Electricals',  '29AAJCB5566H1Z7', 'BE-77',        [20, 10, 2026], 880,     158.4, 0,     0],
];
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const pad = (n) => String(n).padStart(2, '0');
const serial = ([d, m, y]) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
const inr = (n) => n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const book = (rows, name) => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sheet1');
  XLSX.writeFile(wb, path.join(OUT, name));
};

/* ── clean portal file the messy books files are reconciled against ── */
const portalCsv = [
  'GSTIN of supplier,Trade/Legal name,Invoice number,Invoice Date,Taxable Value (₹),Integrated Tax(₹),Central Tax(₹),State/UT Tax(₹)',
  ...TRUTH.map(([n, g, i, d, t, ig, c, s]) => `${g},${n},${i},${pad(d[0])}-${pad(d[1])}-${d[2]},${t.toFixed(2)},${ig.toFixed(2)},${c.toFixed(2)},${s.toFixed(2)}`),
].join('\n');
fs.writeFileSync(path.join(OUT, 'Clean_Portal_GSTR2B_Oct2026.csv'), portalCsv + '\n');

/* ── A. Tally: banner block, text numbers, Dr suffix, 'quote, sci-notation, page-break junk ── */
const tallyHeader = ['Date', 'Particulars', 'Vch Type', 'Vch No.', 'Supplier Invoice No.', 'GSTIN/UIN', 'Gross Total', 'Taxable Value', 'IGST', 'CGST', 'SGST'];
const tallyRow = ([n, g, i, d, t, ig, c, s], k) => [
  `${d[0]}-${MON[d[1] - 1]}-${String(d[2]).slice(2)}`, n, 'Purchase', `PV/${pad(k + 1)}`,
  i === '101000' ? '1.01E+05' : `'${i} `,
  k % 2 ? ` ${g.toLowerCase()} ` : g,
  `${inr(t + ig + c + s)} Dr`, `${inr(t)} Dr`, ig ? inr(ig) : '', c ? inr(c) : '', s ? inr(s) : '',
];
book([
  ['ABC Enterprises Pvt Ltd'], ['Plot 12, MIDC Hingna Road, Nagpur - 440016'], ['GSTIN/UIN: 27ABCDE1234F1Z5'],
  ['Purchase Register'], ['1-Oct-26 to 31-Oct-26'], [],
  tallyHeader,
  ...TRUTH.slice(0, 3).map(tallyRow),
  [], ['', '', '', '', '', '', 'Carried Over', '', '', '', ''], ['Page 2 of 3'],
  tallyHeader,
  ...TRUTH.slice(3).map((r, k) => tallyRow(r, k + 3)),
  [], ['', 'Grand Total', '', '', '', '', '2,07,400.00 Dr', '2,05,080.00 Dr', '2,700.00', '16,800.00', '16,800.00'],
], 'Tally_Messy_Purchase_Register.xlsx');

/* ── B. Zoho Books CSV: UTF-8 BOM, "Bill#", SubTotal, ₹ strings, 2026/10/03 dates, Total footer ── */
const q = (v) => `"${v}"`;
fs.writeFileSync(path.join(OUT, 'Zoho_Bills_Messy.csv'), '\ufeff' + [
  'Bill Date,Bill#,Vendor Name,GSTIN,SubTotal,CGST,SGST,IGST',
  ...TRUTH.map(([n, g, i, d, t, ig, c, s]) => [`${d[2]}/${pad(d[1])}/${pad(d[0])}`, i, n, g, q(`₹${inr(t)}`), q(`₹${inr(c)}`), q(`₹${inr(s)}`), q(`₹${inr(ig)}`)].join(',')),
  `,,Total,,${q('₹2,05,080.00')},${q('₹16,800.00')},${q('₹16,800.00')},${q('₹2,700.00')}`,
].join('\n') + '\n');

/* ── C. SAP: data starts in column C, serial dates, numeric invoice, 'Doc No' + 'Taxable Amt' headers ── */
book([
  [], ['', '', 'Report: ZMIRO_VENDOR_INVOICE_REGISTER'], ['', '', 'Run by: USER01   Client: 100'], [],
  ['', '', 'Posting Date', 'Doc No', 'Vendor', 'Vendor GSTIN', 'Taxable Amt', 'IGST Amt', 'CGST Amt', 'SGST Amt', 'Currency'],
  ...TRUTH.map(([n, g, i, d, t, ig, c, s]) => ['', '', serial(d), i === '101000' ? 101000 : i, n, g, t, ig, c, s, 'INR']),
  ['', '', '', 'Total', '', '', 205080, 2700, 16800, 16800, ''],
], 'SAP_Export_Messy.xlsx');

/* ── D. Hand-made client file: typo'd header, odd names, trailing-space GSTIN, TOTAL row ── */
book([
  [], ['XYZ & Co. - purchase details for October 2026'], [],
  ['S.No', 'Inv Date', 'Party', 'Party GST No', 'Invioce No.', 'Basic Value', 'IGST', 'CGST', 'SGST', 'Total Tax'],
  ...TRUTH.map(([n, g, i, d, t, ig, c, s], k) => [k + 1, `${pad(d[0])}.${pad(d[1])}.${d[2]}`, n, `${g}  `, i, t, ig, c, s, ig + c + s]),
  ['', '', 'TOTAL', '', '', 205080, 2700, 16800, 16800, 36300],
], 'Manual_Client_File_Messy.xlsx');

/* ── E. Unmappable headers → must trigger the manual-mapping modal ── */
fs.writeFileSync(path.join(OUT, 'Needs_Manual_Mapping.csv'), [
  'Sr,Seller Reg,Doc Ref,Memo,Net,Tax Paid',
  ...TRUTH.map(([n, g, i, d, t, ig, c, s], k) => [k + 1, g, i, n, t, ig + c + s].join(',')),
].join('\n') + '\n');

console.log('Messy pack written to', OUT, fs.readdirSync(OUT));
