/**
 * P2.3 — the GST recon output schema is FROZEN. The colleague's IMS list (C6.3),
 * supplier-mismatch list (C6.5) and annual recon (C6.6) read these exact names.
 * If this test fails you changed the contract: update RECON_OUTPUT_SCHEMA.md and tell them.
 */
import { gstMatch } from '../domain-tools/gst/gst-match.tool';
import { normalizeInvoiceNumber } from '../domain-tools/gst/normalize-invoice';
import { NormalizedInvoiceRow } from '../domain-tools/types/domain-tool.types';
import { mapResultRows, mapToSheetActions } from './gst-recon-to-actions.mapper';
import { gstinCheckLabel, sec164Status } from './recon-flags';

const row = (inv: string, extra: Partial<NormalizedInvoiceRow> = {}): NormalizedInvoiceRow => ({
  gstin: '27AABCA1234F1Z9',
  invoiceNumber: inv,
  normalizedInvoiceNumber: normalizeInvoiceNumber(inv),
  invoiceDate: '2026-04-15',
  taxableValue: 10000,
  taxAmount: 1800,
  igst: 0,
  cgst: 900,
  sgst: 900,
  narration: 'ABC Traders',
  documentType: 'invoice',
  sourceRowRef: { documentType: 'workbook', documentId: 'PR', rowOrLine: 2 },
  ...extra,
});
const portal = (inv: string, extra: Partial<NormalizedInvoiceRow> = {}) =>
  row(inv, { sourceRowRef: { documentType: 'gstr2b', documentId: '2B', rowOrLine: 5 }, ...extra });

export const FROZEN_STATUSES = ['MATCHED', 'PARTIAL', 'PR_ONLY', 'PORTAL_ONLY'];
export const FROZEN_ROW_FIELDS = [
  'pr_ref', 'portal_ref', 'status', 'pass', 'confidence', 'difference', 'diff_type', 'itc_amount',
  'ims_status', 'rcm_flag', 'invoice_number', 'gstin', 'vendor_name', 'mismatch_reason', 'explanation',
  'books_sheet_name', 'books_row', 'gstin_check', 'invoice_key', 'sec_16_4',
];
export const FROZEN_DETAIL_HEADERS = [
  'Status', 'Pass', 'GSTIN', 'Vendor Name', 'Invoice No', 'ITC Amount', 'Confidence', 'Difference',
  'Diff Type', 'IMS Status', 'RCM', 'Books Ref', 'Portal Ref', 'GSTIN Check', 'Invoice Key', 'Sec 16(4)',
];

describe('P2.3 frozen GST recon output schema', () => {
  const result = gstMatch({ purchaseRegister: [row('INV/26-27/0233')], gstr2b: [portal('233')] });
  const rows = mapResultRows(result.data.resultRows, new Date(Date.UTC(2026, 10, 7)));

  it('row fields are exactly the frozen list', () => {
    expect(Object.keys(rows[0]).sort()).toEqual([...FROZEN_ROW_FIELDS].sort());
  });

  it('statuses stay inside the frozen bucket names', () => {
    const all = gstMatch({
      purchaseRegister: [row('A-1'), row('ONLY-PR', { gstin: '29AAACM5678K1ZV' })],
      gstr2b: [portal('A-1'), portal('ONLY-2B', { gstin: '07BBBPS4321Q1ZA' })],
    }).data.resultRows.map((r) => r.status);
    for (const s of all) expect(FROZEN_STATUSES).toContain(s);
  });

  it('detail headers are exactly the frozen list', () => {
    const actions = mapToSheetActions({
      sheetName: 'Recon', reconType: 'purchase_vs_2b', runAt: '07-11-2026', relativeTo: 'PR',
      summary: {} as never, rows,
    });
    const table = actions.find((a) => a.type === 'WRITE_TABLE') as { rows: unknown[][] };
    const headerRow = table.rows.find((r) => r[0] === 'Status' && r[1] === 'Pass');
    expect(headerRow).toEqual(FROZEN_DETAIL_HEADERS);
  });

  it('flags a typo GSTIN in the output (never corrects it)', () => {
    const typo = gstMatch({ purchaseRegister: [row('X-1', { gstin: '27AABCA1234F1Z7' })], gstr2b: [] });
    const out = mapResultRows(typo.data.resultRows);
    expect(out[0].gstin).toBe('27AABCA1234F1Z7');
    expect(out[0].gstin_check).toBe('Checksum fail');
    expect(gstinCheckLabel('27AABCA1234F1Z9')).toBe('Valid');
    expect(gstinCheckLabel('')).toBeNull();
  });

  it('invoice key explains the match', () => {
    expect(rows[0].invoice_key).toBe('233');
  });

  it('₹ tolerance moves a ₹50-difference row between buckets', () => {
    const books = row('T-1', { taxableValue: 10000, taxAmount: 1800 });
    const gst = portal('T-1', { taxableValue: 10050, taxAmount: 1809, cgst: 904.5, sgst: 904.5 });
    const tight = gstMatch({ purchaseRegister: [books], gstr2b: [gst], settings: { amountToleranceAbs: 1, amountTolerancePct: 0 } });
    const loose = gstMatch({ purchaseRegister: [books], gstr2b: [gst], settings: { amountToleranceAbs: 100, amountTolerancePct: 0 } });
    const statusOf = (r: typeof tight) => r.data.resultRows.find((x) => x.registerRow)?.status;
    expect(statusOf(tight)).not.toBe(statusOf(loose));
    expect(statusOf(tight)).toBe('PR_ONLY'); // amount mismatch
    expect(statusOf(loose)).toBe('PARTIAL'); // matched within tolerance, difference shown
  });
});

describe('P4.3 Sec 16(4) flag (rule table, CA to verify)', () => {
  it.each([
    ['2025-06-10', '2026-11-07', 'Near', '30-11-2026'],
    ['2025-06-10', '2026-12-05', 'Beyond', '30-11-2026'],
    ['2026-04-15', '2026-11-07', 'Within', '30-11-2027'],
    ['2026-03-31', '2026-10-01', 'Within', '30-11-2026'],
  ])('invoice %s on %s → %s (by %s)', (inv, asOn, status, deadline) => {
    expect(sec164Status(inv, new Date(asOn))).toEqual({ status, deadline });
  });
  it('unreadable date → no flag', () => expect(sec164Status('n/a')).toBeNull());
});
