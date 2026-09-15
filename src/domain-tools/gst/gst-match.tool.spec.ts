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
    const result = gstMatch({
      purchaseRegister: [sampleRow('ONLY-PR')],
      gstr2b: [sampleRow('ONLY-2B', { sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 4 } })],
    });
    expect(result.data.resultRows.some((r) => r.status === 'PR_ONLY')).toBe(true);
    expect(result.data.resultRows.some((r) => r.status === 'PORTAL_ONLY')).toBe(true);
  });
});
