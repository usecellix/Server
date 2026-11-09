import { gstMatch } from './gst-match.tool';
import { gstr1aCorrections } from './irn-pass';
import { normalizeInvoiceNumber } from './normalize-invoice';
import { NormalizedInvoiceRow } from '../types/domain-tool.types';
import { mapResultRows, mapToSheetActions } from '../../gst-recon/gst-recon-to-actions.mapper';

const row = (inv: string, extra: Partial<NormalizedInvoiceRow> = {}): NormalizedInvoiceRow => ({
  gstin: '27AABCA1234F1Z9',
  invoiceNumber: inv,
  normalizedInvoiceNumber: normalizeInvoiceNumber(inv),
  invoiceDate: '2026-05-10',
  taxableValue: 10000,
  taxAmount: 1800,
  igst: 1800,
  cgst: 0,
  sgst: 0,
  narration: 'Customer A',
  documentType: 'invoice',
  sourceRowRef: { documentType: 'workbook', documentId: 'Sales', rowOrLine: 2 },
  ...extra,
});

describe('P3.7 IRN-first matching + GSTR-1A list', () => {
  const IRN = 'a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6abcd';

  it('IRN match beats a different invoice number', () => {
    const r = gstMatch({ mode: 'sales', purchaseRegister: [row('S-0099', { irn: IRN })], gstr2b: [row('SINV/99', { irn: IRN.toUpperCase() })] });
    const hit = r.data.resultRows[0];
    expect(hit.status).toBe('MATCHED');
    expect(hit.pass).toBe(0);
    expect(hit.difference).toMatch(/Matched on IRN; invoice no\. differs/);
  });

  it('same IRN but different amount is PARTIAL with the difference, not silently matched', () => {
    const r = gstMatch({ mode: 'sales', purchaseRegister: [row('S-1', { irn: IRN })], gstr2b: [row('S-1', { irn: IRN, taxableValue: 10500, taxAmount: 1890 })] });
    expect(r.data.resultRows[0]).toMatchObject({ status: 'PARTIAL', diffType: 'IRN_AMOUNT' });
  });

  it('rows without IRN behave exactly as before', () => {
    const r = gstMatch({ mode: 'sales', purchaseRegister: [row('S-2')], gstr2b: [row('S-2')] });
    expect(r.data.resultRows[0]).toMatchObject({ status: 'MATCHED', pass: 1 });
  });

  it('GSTR-1A list contains only mismatches', () => {
    const r = gstMatch({
      mode: 'sales',
      purchaseRegister: [row('S-1'), row('S-2', { taxableValue: 9000, taxAmount: 1620 }), row('S-3', { gstin: '29AAACM5678K1ZV' })],
      gstr2b: [row('S-1'), row('S-2')],
    });
    const list = gstr1aCorrections(r.data.resultRows);
    expect(list.every((x) => x.row.status !== 'MATCHED')).toBe(true);
    expect(list.length).toBeGreaterThanOrEqual(2);
  });

  it('the sales report sheet carries a GSTR-1A CORRECTION LIST section', () => {
    const r = gstMatch({ mode: 'sales', purchaseRegister: [row('S-1'), row('S-9', { gstin: '29AAACM5678K1ZV' })], gstr2b: [row('S-1')] });
    const actions = mapToSheetActions({ sheetName: 'Recon', reconType: 'sales_vs_gstr1', runAt: '07-11-2026', summary: {} as never, rows: mapResultRows(r.data.resultRows), isSales: true });
    const table = actions.find((a) => a.type === 'WRITE_TABLE') as { rows: unknown[][] };
    expect(table.rows.some((x) => String(x[0]).startsWith('GSTR-1A CORRECTION LIST'))).toBe(true);
    expect(table.rows.some((x) => x[0] === 'Report missing invoice in GSTR-1A')).toBe(true);
  });
});
