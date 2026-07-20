import { gstMatch } from './gst-match.tool';
import { expectStubNotImplemented } from '../test-utils/domain-tool-test.util';
import { NormalizedInvoiceRow } from '../types/domain-tool.types';

const sampleRow = (invoiceNumber: string): NormalizedInvoiceRow => ({
  gstin: '32AAAAA0000A1Z5',
  invoiceNumber,
  invoiceDate: '2026-04-15',
  taxableValue: 10000,
  taxAmount: 1800,
  sourceRowRef: {
    documentType: 'workbook',
    documentId: 'synth-register',
    rowOrLine: 2,
  },
});

describe('gstMatch', () => {
  it('has DomainTool signature and throws Not implemented (stub)', () => {
    expectStubNotImplemented(gstMatch, {
      purchaseRegister: [sampleRow('INV-1')],
      gstr2b: [sampleRow('INV-1')],
      matchKeys: ['gstin', 'invoiceNumber'],
      amountTolerance: 1,
    });
  });
});
