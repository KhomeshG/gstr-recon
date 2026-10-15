'use strict';
// Every messy export must reconcile to the SAME result. Run: node test-data/verify-messy.js
const fs = require('fs');
const path = require('path');
const { processUpload } = require('../utils/reconcile');

const dir = path.join(__dirname, 'messy');
const file = (name) => ({ buffer: fs.readFileSync(path.join(dir, name)), originalname: name });
const portal = file('Clean_Portal_GSTR2B_Oct2026.csv');

// expectations per file: header rows above table, footer/junk rows removed, columns recognised
const CASES = [
  { name: 'Tally_Messy_Purchase_Register.xlsx', top: 6, removed: 4, note: false },
  { name: 'Zoho_Bills_Messy.csv',               top: 0, removed: 1, note: true  },   // SubTotal is a weak guess → note
  { name: 'SAP_Export_Messy.xlsx',              top: 4, removed: 1, note: true  },   // "Doc No" is a weak guess → note
  { name: 'Manual_Client_File_Messy.xlsx',      top: 3, removed: 1, note: true  },   // typo "Invioce No." → fuzzy note
];

let bad = 0;
const check = (ok, msg) => { if (!ok) bad++; console.log(`${ok ? 'PASS' : 'FAIL'}  ${msg}`); };

CASES.forEach(({ name, top, removed, note }) => {
  const out = processUpload(file(name), portal, { period: '2026-10' });
  console.log(`\n── ${name}`);
  if (!out.ok) return check(false, `parsed (${out.error || 'needs manual mapping'})`);
  const c = out.cleaning[0];
  console.log(`   ${c.text}`);
  c.notes.forEach((n) => console.log(`   note: ${n}`));
  check(out.results.exact.length === 6, `6 exact matches (got ${out.results.exact.length})`);
  check(out.inputs.skippedRows.length === 0, `0 skipped rows (got ${out.inputs.skippedRows.length})`);
  check(out.results.missingInPortal.length === 0 && out.results.missingInBooks.length === 0, 'nothing missing on either side');
  check(c.stats.topRowsSkipped === top, `${top} top rows skipped (got ${c.stats.topRowsSkipped})`);
  check(c.stats.removedRows === removed, `${removed} junk row(s) removed (got ${c.stats.removedRows})`);
  check(Boolean(c.notes.length) === note, `weak-match note ${note ? 'shown' : 'absent'}`);
  const rec = out.inputs.books.find((r) => r.invoice === 'INV2026101' || r.gstin === '27AABCS1234A1Z5');
  check(/^2026-10-03$/.test(rec.date), `date standardised to YYYY-MM-DD (got "${rec.date}")`);
  check(out.inputs.books.some((r) => r.invoiceRaw === '101000'), 'sci-notation / numeric invoice → "101000"');
  check(!out.inputs.books.some((r) => /^'/.test(r.invoiceRaw) || /\s$/.test(r.invoiceRaw)), 'no leading quote / trailing space left in invoice numbers');
});

console.log('\n── Needs_Manual_Mapping.csv');
const first = processUpload(file('Needs_Manual_Mapping.csv'), portal, { period: '2026-10' });
check(first.ok === false && first.needsManualMapping === true, 'flags needsManualMapping instead of crashing');
const req = first.requests?.[0];
check(Array.isArray(first.detectedHeaders) && first.detectedHeaders.length === 6, `JSON has detectedHeaders (${first.detectedHeaders})`);
check(Array.isArray(first.rawPreviewRows) && first.rawPreviewRows.length >= 7, `JSON has rawPreviewRows (${first.rawPreviewRows?.length} rows)`);
check(req && req.missingFields.join() === 'gstin,invoice,taxable', `all three mandatory fields reported missing (${req?.missingFields})`);

const mapped = processUpload(file('Needs_Manual_Mapping.csv'), portal, {
  period: '2026-10',
  overrides: { books: { headerRowIndex: 0, columns: { gstin: 1, invoice: 2, taxable: 4, name: 3, totalTax: 5, date: null, igst: null, cgst: null, sgst: null } } },
});
check(mapped.ok && mapped.results.exact.length === 6, `after 1-click mapping → 6 exact matches (ok=${mapped.ok})`);
check(mapped.ok && mapped.cleaning[0].manual === true, 'badge marks the columns as manually mapped');

console.log(bad ? `\n${bad} check(s) failed` : '\nAll messy-export checks passed ✔');
process.exit(bad ? 1 : 0);
