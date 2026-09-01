import { itcCompute } from './itc-compute.tool';
import { assertDomainToolResultShape } from '../test-utils/domain-tool-test.util';
import { NormalizedInvoiceRow } from '../types/domain-tool.types';

const row = (tax: number): NormalizedInvoiceRow => ({
  gstin: '32AAAAA0000A1Z5',
  invoiceNumber: 'INV-1',
  normalizedInvoiceNumber: 'INV1',
  invoiceDate: '2026-04-01',
  taxableValue: 10000,
  taxAmount: tax,
  igst: 0,
  cgst: tax / 2,
  sgst: tax / 2,
  narration: '',
  documentType: 'invoice',
  sourceRowRef: { documentType: 'purchase_register', documentId: 'pr', rowOrLine: 2 },
});

describe('itcCompute', () => {
  it('computes total from eligible invoices', () => {
    const result = itcCompute({
      eligibleInvoices: [row(1800), row(900)],
    });
    assertDomainToolResultShape(result);
    expect(result.data.totalItcClaimable).toBe(2700);
    expect(result.data.totalMatched).toBe(2);
  });
});
