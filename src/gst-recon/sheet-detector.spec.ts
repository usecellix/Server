import {
  missingSheetChatMessage,
  pickBestSheet,
  scoreSheetRole,
} from './sheet-detector';

describe('sheet-detector', () => {
  it('scores GSTR-2B portal headers highly', () => {
    const score = scoreSheetRole(
      'Portal Export',
      ['GSTIN of supplier', 'Invoice number', 'Taxable Value', 'Document Type'],
      'GSTR2B',
    );
    expect(score).toBeGreaterThanOrEqual(3);
  });

  it('scores purchase register by GSTIN + invoice + taxable', () => {
    const score = scoreSheetRole(
      'Books',
      ['Supplier GSTIN', 'Invoice No', 'Taxable Amount', 'CGST', 'SGST'],
      'PURCHASE_REGISTER',
    );
    expect(score).toBeGreaterThanOrEqual(3);
  });

  it('picks unique best sheet', () => {
    const result = pickBestSheet(
      [
        {
          sheetName: 'GSTR-2B Apr',
          headers: ['GSTIN of supplier', 'Invoice number', 'Taxable Value'],
        },
        {
          sheetName: 'Notes',
          headers: ['Comment'],
        },
      ],
      'GSTR2B',
    );
    expect(result.best?.sheetName).toBe('GSTR-2B Apr');
    expect(result.ties).toEqual([]);
  });

  it('builds missing GSTR-2B chat message', () => {
    const msg = missingSheetChatMessage(['GSTR2B']);
    expect(msg).toMatch(/GSTR-2B/);
    expect(msg).toMatch(/GST portal/i);
  });
});
