import { normalizeInvoiceNumber, stringSimilarityPercent } from './normalize-invoice';
import { gstMatch } from './gst-match.tool';
import { assertDomainToolResultShape } from '../test-utils/domain-tool-test.util';
import { NormalizedInvoiceRow } from '../types/domain-tool.types';

const sampleRow = (
  invoiceNumber: string,
  extra: Partial<NormalizedInvoiceRow> = {},
): NormalizedInvoiceRow => ({
  gstin: '32AAAAA0000A1Z5',
  invoiceNumber,
  normalizedInvoiceNumber: normalizeInvoiceNumber(invoiceNumber),
  invoiceDate: '2026-04-15',
  taxableValue: 10000,
  taxAmount: 1800,
  igst: 0,
  cgst: 900,
  sgst: 900,
  narration: '',
  documentType: 'invoice',
  sourceRowRef: {
    documentType: 'workbook',
    documentId: 'synth-register',
    rowOrLine: 2,
  },
  ...extra,
});

describe('normalize-invoice', () => {
  it('normalizes invoice separators', () => {
    expect(normalizeInvoiceNumber('INV/1001')).toBe(normalizeInvoiceNumber('INV-1001'));
  });

  it('scores fuzzy similarity', () => {
    expect(stringSimilarityPercent('INV1001', 'INV1001')).toBe(100);
  });
});

describe('gstMatch', () => {
  it('exact-matches identical invoices', () => {
    const result = gstMatch({
      purchaseRegister: [sampleRow('INV-1')],
      gstr2b: [sampleRow('INV-1', { sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 3 } })],
      matchKeys: ['gstin', 'invoiceNumber'],
      amountTolerance: 1,
    });
    assertDomainToolResultShape(result);
    expect(result.data.matched.length).toBe(1);
    expect(result.data.resultRows[0].status).toBe('MATCHED');
  });

  it('marks books-only as PR_ONLY', () => {
    // Different GSTINs — genuinely unrelated invoices, not a same-GSTIN near-miss
    // pair. (Sharing a GSTIN here, even with a different invoice number, made the
    // portal row look like a date+amount "closest candidate" for the books row's
    // diagnosis, which — correctly, per BUGFIX_portal_only_double_count.md — now
    // removes a cited candidate from Portal-only. That's the right behavior for a
    // real near-miss; it just wasn't this test's intent.)
    const result = gstMatch({
      purchaseRegister: [sampleRow('ONLY-PR')],
      gstr2b: [
        sampleRow('ONLY-2B', {
          gstin: '27ZZZZZ0000Z1Z5',
          sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 4 },
        }),
      ],
    });
    expect(result.data.resultRows.some((r) => r.status === 'PR_ONLY')).toBe(true);
    expect(result.data.resultRows.some((r) => r.status === 'PORTAL_ONLY')).toBe(true);
  });
});

describe('gstMatch — RCM detection (keyword-gated, never on blank GSTIN alone)', () => {
  it('does NOT flag a blank-GSTIN row as RCM when narration has no RCM keyword — it must fall through to blank_counterparty_gstin diagnosis instead', () => {
    // Regression guard: an earlier version of runPassRcm treated blank/invalid GSTIN
    // alone as sufficient, which silently stole every blank-GSTIN row away from the
    // blank_counterparty_gstin diagnosis this session fixed to work correctly. A blank
    // GSTIN by itself is only *potentially* RCM-eligible, not automatically RCM.
    const result = gstMatch({
      purchaseRegister: [sampleRow('', { gstin: '', narration: 'Office stationery purchase' })],
      gstr2b: [],
      settings: { detectRcm: true },
    });
    const row = result.data.resultRows[0];
    expect(row.status).not.toBe('RCM');
    expect(row.status).toBe('PR_ONLY');
    expect(row.mismatchReason).toBe('blank_counterparty_gstin');
  });

  it('flags a row with a valid GSTIN as RCM when the narration hits an RCM keyword', () => {
    const result = gstMatch({
      purchaseRegister: [
        sampleRow('FRT-1', { narration: 'Freight charges paid to GTA for material transport' }),
      ],
      gstr2b: [],
      settings: { detectRcm: true },
    });
    const row = result.data.resultRows[0];
    expect(row.status).toBe('RCM');
    expect(row.rcmFlag).toBe(true);
    expect(row.confidence).toBe(0.7);
  });

  it('flags a blank-GSTIN row as RCM (higher confidence) when it ALSO hits an RCM keyword', () => {
    const result = gstMatch({
      purchaseRegister: [
        sampleRow('', { gstin: '', narration: 'Legal fee paid to advocate for GST case' }),
      ],
      gstr2b: [],
      settings: { detectRcm: true },
    });
    const row = result.data.resultRows[0];
    expect(row.status).toBe('RCM');
    expect(row.confidence).toBe(0.85);
  });

  it('never runs RCM detection when detectRcm is false', () => {
    const result = gstMatch({
      purchaseRegister: [sampleRow('FRT-1', { narration: 'Freight charges via GTA' })],
      gstr2b: [],
      settings: { detectRcm: false },
    });
    expect(result.data.resultRows[0].status).not.toBe('RCM');
  });
});

describe('gstMatch — Amended invoices (doc type "IA" / narration AMEND) get their own category, not CREDIT_NOTE', () => {
  it('produces AMENDED, not CREDIT_NOTE, when the portal counterpart is an amendment', () => {
    const pr = sampleRow('INV-AMD-1', {
      documentType: 'credit_note',
      taxableValue: -5000,
      narration: 'Credit note for INV-AMD-1',
    });
    const portal = sampleRow('IA-AMD-1', {
      documentType: 'amended',
      taxableValue: -5000,
      narration: 'Amended invoice IA reference',
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 9 },
    });
    const result = gstMatch({
      purchaseRegister: [pr],
      gstr2b: [portal],
    });
    const row = result.data.resultRows.find((r) => r.registerRow === pr);
    expect(row?.status).toBe('AMENDED');
    expect(row?.status).not.toBe('CREDIT_NOTE');
    expect(row?.difference).toMatch(/amended invoice/i);
    expect(row?.difference).toMatch(/verify against the original/i);
  });

  it('still produces CREDIT_NOTE for a genuine credit/debit note pair (unaffected by the amended split)', () => {
    // Different invoice numbers on each side (a books-side CN reference vs. the
    // portal's own CN reference) — same shape as a real credit note pair, and
    // deliberately NOT identical, so Pass 1's exact GSTIN+invoice-number match doesn't
    // claim it before runPassCdn gets a chance to.
    const pr = sampleRow('CN-BOOKS-1', { documentType: 'credit_note', taxableValue: -2000 });
    const portal = sampleRow('CN-PORTAL-1', {
      documentType: 'credit_note',
      taxableValue: -2000,
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 10 },
    });
    const result = gstMatch({
      purchaseRegister: [pr],
      gstr2b: [portal],
    });
    const row = result.data.resultRows.find((r) => r.registerRow === pr);
    expect(row?.status).toBe('CREDIT_NOTE');
  });
});
