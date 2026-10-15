'use strict';
// Unit tests for utils/normalizer.js — run: node test-data/test-normalizer.js
const assert = require('assert');
const N = require('../utils/normalizer');

let pass = 0; const fails = [];
const t = (name, fn) => { try { fn(); pass++; } catch (e) { fails.push(`${name}\n     ${e.message.split('\n')[0]}`); } };
const eq = (a, b) => assert.deepStrictEqual(a, b);

/* ── Invoice numbers ── */
t('sci notation text → 101000', () => eq(N.normalizeInvoiceNumber('1.01E+05'), '101000'));
t('sci notation lower-case', () => eq(N.normalizeInvoiceNumber('1.5e3'), '1500'));
t('numeric cell → text', () => eq(N.normalizeInvoiceNumber(1045), '1045'));
t('leading apostrophe stripped', () => eq(N.normalizeInvoiceNumber("'INV-0045"), 'INV-0045'));
t('curly apostrophe stripped', () => eq(N.normalizeInvoiceNumber('’00123'), '00123'));
t('trailing spaces trimmed', () => eq(N.normalizeInvoiceNumber('  INV/2026/01   '), 'INV/2026/01'));
t('internal slashes+dashes kept', () => eq(N.normalizeInvoiceNumber('AB-12/26'), 'AB-12/26'));
t('nbsp / double spaces collapsed', () => eq(N.normalizeInvoiceNumber('INV\u00a0  7'), 'INV 7'));
t('leading zeros kept', () => eq(N.normalizeInvoiceNumber('000789'), '000789'));
t('blank → empty', () => eq(N.normalizeInvoiceNumber(undefined), ''));

/* ── Amounts ── */
const amt = (v) => N.parseAmount(v);
t('₹ and commas', () => eq(amt('₹ 1,25,000.00'), { value: 125000, status: 'ok' }));
t('$ symbol', () => eq(amt('$1,234.567'), { value: 1234.57, status: 'ok' }));
t('Rs. prefix', () => eq(amt('Rs. 500'), { value: 500, status: 'ok' }));
t('parentheses = negative', () => eq(amt('(250.50)'), { value: -250.5, status: 'ok' }));
t('leading minus', () => eq(amt('-2,500'), { value: -2500, status: 'ok' }));
t('SAP trailing minus', () => eq(amt('1,500.00-'), { value: -1500, status: 'ok' }));
t('Tally Dr suffix', () => eq(amt('11,800.00 Dr'), { value: 11800, status: 'ok' }));
t('Tally Cr suffix', () => eq(amt('11,800.00Cr'), { value: 11800, status: 'ok' }));
t('rounds to 2 dp', () => eq(amt(15000.004), { value: 15000, status: 'ok' }));
t('number cell', () => eq(amt(99.999), { value: 100, status: 'ok' }));
t('blank cell', () => eq(amt('').status, 'blank'));
t('dash placeholder = blank', () => eq(amt('-').status, 'blank'));
t('NIL = blank', () => eq(amt('Nil').status, 'blank'));
t('words are invalid', () => eq(amt('abc').status, 'invalid'));
t('mixed junk is invalid', () => eq(amt('12abc').status, 'invalid'));
t('scientific amount', () => eq(amt('1.5E+3'), { value: 1500, status: 'ok' }));
t('NaN is invalid', () => eq(amt(NaN).status, 'invalid'));

/* ── Dates ── */
const d = N.normalizeDate;
t('serial 45201 → 2023-10-02', () => eq(d(45201), '2023-10-02'));
t('dd-mm-yyyy', () => eq(d('02-10-2026'), '2026-10-02'));
t('yyyy/mm/dd', () => eq(d('2026/10/02'), '2026-10-02'));
t('d-Mon-yy', () => eq(d('2-Oct-26'), '2026-10-02'));
t('dd/mm/yyyy slash', () => eq(d('02/10/2026'), '2026-10-02'));
t('dd.mm.yyyy dots', () => eq(d('02.10.2026'), '2026-10-02'));
t('dd Month yyyy', () => eq(d('2 October 2026'), '2026-10-02'));
t('Mon d, yyyy', () => eq(d('Oct 2, 2026'), '2026-10-02'));
t('dd-mm-yy', () => eq(d('02-10-26'), '2026-10-02'));
t('ISO with time', () => eq(d('2026-10-02 00:00:00'), '2026-10-02'));
t('Sept spelling', () => eq(d('5-Sept-2026'), '2026-09-05'));
t('month-first fallback 10/25/2026', () => eq(d('10/25/2026'), '2026-10-25'));
t('impossible date → blank', () => eq(d('31-02-2026'), ''));
t('Date object', () => eq(d(new Date(2026, 9, 2)), '2026-10-02'));
t('numeric string serial', () => eq(d('45201'), '2023-10-02'));
t('garbage → blank', () => eq(d('soon'), ''));
t('small number is not a serial', () => eq(d(12), ''));

/* ── GSTIN ── */
t('gstin cleaned', () => eq(N.normalizeGSTIN(" 27 aabcs1234a1z5 '"), '27AABCS1234A1Z5'));

/* ── Header mapping: spec aliases ── */
const m = (...h) => N.mapColumns(h);
t('Tally: GSTIN/UIN', () => eq(m('GSTIN/UIN').gstin.index, 0));
t('SAP: Vendor GSTIN', () => eq(m('Vendor GSTIN').gstin.confidence, 1));
t('"GST No" → gstin', () => eq(m('GST No').gstin.index, 0));
t('"Inv No." → invoice', () => eq(m('Inv No.').invoice.index, 0));
t('"Bill #" → invoice', () => eq(m('Bill #').invoice.index, 0));
t('"Vch No" → invoice (weak)', () => { const x = m('Vch No').invoice; eq(x.index, 0); assert(x.confidence < 1); });
t('Supplier Invoice No beats Voucher No', () => eq(m('Voucher No.', 'Supplier Invoice No.').invoice.index, 1));
t('Vch Date → date', () => eq(m('Vch Date').date.index, 0));
t('Invoice Date beats Vch Date', () => eq(m('Vch Date', 'Invoice Date').date.index, 1));
t('"Assessable Value" → taxable', () => eq(m('Assessable Value').taxable.index, 0));
t('"Sub Total" → taxable (weak)', () => eq(m('Sub Total').taxable.index, 0));
t('"Taxable Amount (₹)" → taxable', () => eq(m('Taxable Amount (₹)').taxable.index, 0));
t('bare "Amount" is only low confidence', () => assert(m('Amount').taxable.confidence < N.CONFIDENT));
t('"Integrated Tax" → igst', () => eq(m('Integrated Tax').igst.index, 0));
t('"State/UT Tax(₹)" → sgst', () => eq(m('State/UT Tax(₹)').sgst.index, 0));
t('"SGST/UTGST Amt" → sgst', () => eq(m('SGST/UTGST Amt').sgst.index, 0));
t('"Total Tax (₹)" → totalTax', () => eq(m('Total Tax (₹)').totalTax.index, 0));
t('typo "Invioce No" fuzzy-matches', () => eq(m('Invioce No').invoice.index, 0));
t('"Party Name" → name', () => eq(m('Party Name').name.index, 0));

/* ── Header mapping: traps ── */
t('IGST % column is not the amount', () => eq(m('IGST %', 'IGST Amount').igst.index, 1));
t('"CGST Rate" rejected', () => eq(m('CGST Rate').cgst, undefined));
t('Non-Taxable Value is not taxable', () => eq(m('Non-Taxable Value', 'Taxable Value').taxable.index, 1));
t('Total Taxable Amount ≠ Total Tax', () => eq(m('Total Taxable Amount').totalTax, undefined));
t('Invoice Value (gross) ≠ taxable', () => eq(m('Invoice Value').taxable, undefined));
t('Tax Amount ≠ taxable', () => eq(m('Tax Amount').taxable, undefined));
t('each column used once', () => { const x = m('Amount', 'Amount'); eq(Object.keys(x).length, 1); });

/* ── Row filters ── */
t('GRAND TOTAL row', () => assert(N.isSummaryRow(['', 'GRAND TOTAL', '', 5000], '')));
t('Sub-Total row', () => assert(N.isSummaryRow(['Sub-Total', 100], '')));
t('Carried over', () => assert(N.isSummaryRow(['Carried Over', 100], '')));
t('Page 2 of 3', () => assert(N.isSummaryRow(['Page 2 of 3'], '')));
t('"Total" with colon', () => assert(N.isSummaryRow(['Total :', 1], '')));
t('"Total for Sept" with no GSTIN', () => assert(N.isSummaryRow(['Total for September', 1], '')));
t('supplier named Total Logistics is KEPT', () => eq(N.isSummaryRow(['Total Logistics', 'INV-1'], '27AABCS1234A1Z5'), false));
t('normal data row is not summary', () => eq(N.isSummaryRow(['Shree Traders', 'INV-1', 100], '27AABCS1234A1Z5'), false));
t('repeated header detected', () => assert(N.isRepeatedHeader(['Date', 'Party', 'GSTIN', 'Amount'], ['Date', 'Party', 'GSTIN', 'Amount'])));
t('data row is not a repeated header', () => eq(N.isRepeatedHeader(['2026-10-01', 'X', '27AA', 5], ['Date', 'Party', 'GSTIN', 'Amount']), false));

/* ── Header detection with banners ── */
t('header found under 6 banner rows', () => {
  const rows = [['ABC Enterprises Pvt Ltd'], ['12 MG Road, Nagpur'], ['GSTIN: 27ABCDE1234F1Z5'], ['Purchase Register'], ['1-Sep-26 to 30-Sep-26'], [],
    ['Date', 'Party Name', 'GSTIN/UIN', 'Voucher No', 'Taxable Value'], ['2-Oct-26', 'X', '27AABCS1234A1Z5', '1', 100]];
  const h = N.detectHeader(rows);
  eq(h.startIndex, 6); eq(h.map.gstin.index, 2); eq(h.map.taxable.index, 4);
});
t('two-row merged header (GSTR-2B)', () => {
  const rows = [['GSTIN of supplier', 'Trade/Legal name', 'Invoice details', '', '', 'Taxable Value (₹)'], ['', '', 'Invoice number', 'Invoice type', 'Invoice Date', ''], ['27AABCS1234A1Z5', 'X', 'A1', 'Regular', 45201, 100]];
  const h = N.detectHeader(rows);
  eq(h.startIndex, 0); eq(h.index, 1); eq(h.map.invoice.index, 2);
});

console.log(`${pass} passed, ${fails.length} failed`);
fails.forEach((f) => console.log('  FAIL', f));
process.exit(fails.length ? 1 : 0);
