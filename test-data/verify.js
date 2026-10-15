'use strict';
// Runs the generated files through the real engine and compares with the answer key.
const fs = require('fs');
const path = require('path');
const { processUpload } = require('../utils/reconcile');

const dir = path.join(__dirname, 'files');
const file = (name) => ({ buffer: fs.readFileSync(path.join(dir, name)), originalname: name });
const { expected, entries } = JSON.parse(fs.readFileSync(path.join(dir, 'expected.json'), 'utf8'));

const out = processUpload(file('Purchase_Register_Tally_Style_Sep2026.xlsx'), file('GSTR2B_Official_Format_Sep2026.xlsx'));
if (!out.ok) { console.error('FAILED TO PARSE:', out.error); process.exit(1); }
const { results: r, summary: s } = out;

const actual = {
  booksUsable: s.booksCount, portalUsable: s.portalCount, exact: r.exact.length,
  probable: r.mismatches.filter((m) => !m.reason.startsWith('Taxable')).length,
  value: r.mismatches.filter((m) => m.reason.startsWith('Taxable')).length,
  missingPortal: r.missingInPortal.length, missingBooks: r.missingInBooks.length,
  skippedBooksRows: Number((out.warnings.join(' ').match(/skipped (\d+)/) || [])[1] || 0),
  itcAtRisk: s.itcAtRisk,
};
let bad = 0;
Object.keys(expected).forEach((k) => {
  const ok = actual[k] === expected[k];
  if (!ok) bad++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${k.padEnd(18)} expected ${expected[k]}  got ${actual[k]}`);
});

// per-invoice placement
const where = { B: new Map(), P: new Map() };
const put = (side, inv, bucket) => where[side].set(inv, [...(where[side].get(inv) || []), bucket]);
r.exact.forEach((m) => { put('B', m.books.invoiceRaw, 'EXACT'); put('P', m.portal.invoiceRaw, 'EXACT'); });
r.mismatches.forEach((m) => { const b = m.reason.startsWith('Taxable') ? 'VALUE_MISMATCH' : 'PROBABLE'; put('B', m.books.invoiceRaw, b); put('P', m.portal.invoiceRaw, b); });
r.missingInPortal.forEach((x) => put('B', x.invoiceRaw, 'MISSING_PORTAL'));
r.missingInBooks.forEach((x) => put('P', x.invoiceRaw, 'MISSING_BOOKS'));
entries.filter((e) => e.id !== 'T10a' && e.id !== 'T10b').forEach((e) => {
  const want = e.exp === 'EXACT*' ? 'EXACT' : e.exp;
  [['B', e.b], ['P', e.p]].filter(([, inv]) => inv).forEach(([sd, inv]) => {
    const got = (where[sd].get(inv) || []).join('/') || 'NOT FOUND';
    if (got !== want) { bad++; console.log(`FAIL  ${e.id} ${sd}:${inv} expected ${want} got ${got}`); }
  });
});
console.log(bad ? `\n${bad} check(s) failed` : '\nAll checks passed ✔');
console.log('Warnings:', out.warnings);
process.exit(bad ? 1 : 0);
