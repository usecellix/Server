import { checkEInvoiceRow, checkTdsRow, eInvoiceLane, tdsCheckLane } from './compliance.lanes';
import { cellsOf, makeCtx, serial } from './lane-test-utils';
import { LaneResult } from './lane.types';
import { matchingLanes } from './lane-registry';

// Same rows as T_TDS / T_EInvoice in CELLIX_TEST_DATA_EXTRA.xlsx
export const T_TDS = [
  ['Deductee', 'PAN', 'Section', 'Amount Paid', 'TDS Rate', 'TDS Deducted'],
  ['Sharma Contractors', 'AAAPS1234A', '194C', 50000, 1, 500],
  ['Mehta Builders Pvt Ltd', 'AABCM1234B', '194C', 80000, 1, 800],
  ['Rao & Associates', 'AAAFR1234C', '194J', 120000, 10, 12000],
  ['Kumar Tech Services', '', '194J(a)', 60000, 2, 1200],
  ['City Estates', 'AABCC4321D', '194I(b)', 40000, 10, 4000],
  ['Gupta Commission', 'AAAPG5678E', '194H', 25000, 5, 1250],
  ['Unknown Payee', 'AAAPX1111Z', '195X', 10000, 1, 100],
];

const IRN1 = 'c4cedd40e00558b82d714398828388e16831f3d018acd3924b8c651b09a96f11';
export const T_EINVOICE = [
  ['IRN', 'Ack No', 'Invoice No', 'Invoice Date', 'Supplier GSTIN', 'Buyer GSTIN', 'Taxable Value'],
  [IRN1, '112010012345678', 'INV/26/001', serial(2026, 10, 1), '27AABCA1234F1Z9', '29AAACM5678K1ZV', 100000],
  ['abc123', '112010012345679', 'INV/26/002', serial(2026, 10, 2), '27AABCA1234F1Z9', '29AAACM5678K1ZV', 50000],
  [IRN1, '112010012345680', 'INV/26/003', serial(2026, 10, 3), '27AABCA1234F1Z9', '07BBBPS4321Q1ZA', 20000],
  ['862a631893b64d2f19ce85165107db500cbf1e234f2487c61f7dba0520269980', '112010012345681', '0INV-55', serial(2026, 10, 4), '27AABCA1234F1Z9', '33CCCHD9876L1ZN', 30000],
  ['408d7a657150cf5cc1c364470a5eff90d5d9ed330717242e587a0a725ee085e9', '112010012345682', 'INV/26/005', serial(2026, 10, 5), '27AABCA1234F1Z7', '24DDDFR1111M1ZA', 40000],
  ['818eed4adec9cf2da9deaea5ec169d4e296d13485b32884f3c88348228d50880', '112010012345683', 'INV/26/006', serial(2026, 12, 31), '27AABCA1234F1Z9', '19EEETP2222N1ZY', 60000],
  ['d8303960c9b7682ad3fd0a807ae8de6d167b8f23ef205b5f1374042d471fc1d9', '112010012345684', 'INV/26/007', serial(2026, 10, 7), '27AABCA1234F1Z9', '27AABCA1234F1Z9', 70000],
  ['bfc1125fd2363a32f6aaeaed4616b10216106c428acc372c6ad37eb9e36c7707', '112010012345685', 'INVOICE/2026/0000123', serial(2026, 10, 8), '27AABCA1234F1Z9', '36FFFCT3333P1ZJ', 80000],
];

describe('TDS section / rate checker', () => {
  const res = tdsCheckLane.run(makeCtx(T_TDS, { message: 'Check the TDS sections and rates in this sheet' })) as LaneResult;
  const cells = cellsOf(res);

  it('2 OK, 5 flagged; adds three columns, source columns untouched', () => {
    expect(res.answer).toMatch(/Checked \*\*7\*\* TDS rows: \*\*2 OK\*\*, \*\*5 flagged\*\*/);
    expect(cells.G1).toBe('Expected Rate %');
    expect(cells.I1).toBe('TDS Check');
    expect(Object.keys(cells).some((a) => /^[A-F]\d/.test(a))).toBe(false);
  });

  it.each([
    ['I2', /^OK$/],
    ['I3', /Rate 1% vs 2% for 194C \(others\); Short deducted by ₹800/],
    ['I4', /^OK$/],
    ['I5', /PAN missing — higher rate 20%.*Rate 2% vs 20%.*Short deducted by ₹10,800/],
    ['I6', /Below the 194I\(b\) threshold/],
    ['I7', /Rate 5% vs 2%.*Excess deducted by ₹750/],
    ['I8', /Unknown section "195X"/],
  ])('%s', (cell, re) => expect(String(cells[cell])).toMatch(re));

  it('cites the rule table and adds the disclaimer', () => {
    expect(res.answer).toMatch(/tds-sections vFY2026-27\.v1 \(CA to verify\)/);
    expect(res.answer).toMatch(/decision rests with you/);
  });

  it('aggregate threshold: two 194C payments of 25,000 each to the same PAN cross ₹1 lakh only above 1 lakh', () => {
    expect(checkTdsRow({ section: '194C', pan: 'AAAPS1234A', amount: 25000, tds: 0 }, 50000).issues).toEqual([]);
    expect(checkTdsRow({ section: '194C', pan: 'AAAPS1234A', amount: 25000, tds: 0 }, 120000).issues).toEqual(['Short deducted by ₹250']);
  });

  it('routing: not the TDS template, dropdown or 26AS match', () => {
    expect(matchingLanes('Check the TDS sections and rates in this sheet', true).map((l) => l.id)).toContain('tds-check');
    for (const m of ['Create a TDS working template', 'Add a dropdown of TDS sections to column C', 'Match TDS with 26AS']) {
      expect(matchingLanes(m, true).map((l) => l.id)).not.toContain('tds-check');
    }
  });
});

describe('e-Invoice (IRN) validation', () => {
  const res = eInvoiceLane.run(makeCtx(T_EINVOICE, { message: 'Validate the e-invoices in this sheet' })) as LaneResult;
  const cells = cellsOf(res);

  it('1 OK, 7 to check', () => {
    expect(res.answer).toMatch(/Validated \*\*8\*\* e-invoices: \*\*1 OK\*\*, \*\*7 to check\*\*/);
    expect(cells.H1).toBe('e-Invoice Check');
    expect(cells.H2).toBe('OK');
  });

  it.each([
    ['I3', /IRN must be 64 hex characters \(has 6\)/],
    ['I4', /Duplicate IRN \(same as row 2\)/],
    ['I5', /cannot start with 0/],
    ['I6', /Supplier GSTIN: Checksum fail/],
    ['I7', /Invoice date is in the future/],
    ['I8', /Buyer GSTIN equals supplier GSTIN/],
    ['I9', /longer than 16 characters/],
  ])('%s', (cell, re) => expect(String(cells[cell])).toMatch(re));

  it('states it is a format check only and cites the rule', () => {
    expect(res.answer).toMatch(/cannot confirm an IRN is registered/);
    expect(res.answer).toMatch(/e-invoice v2026-11\.v1/);
  });

  it('URP buyer (unregistered) is accepted', () => {
    expect(checkEInvoiceRow({ buyerGstin: 'URP', sellerGstin: '27AABCA1234F1Z9' }, new Date())).toEqual([]);
  });

  it('routing', () => {
    expect(matchingLanes('Validate the e-invoices in this sheet', true).map((l) => l.id)).toContain('e-invoice');
    expect(matchingLanes('Check the IRNs', true).map((l) => l.id)).toContain('e-invoice');
    expect(matchingLanes('Reconcile sales register with GSTR-1', true).map((l) => l.id)).not.toContain('e-invoice');
  });
});
