import {
  ambiguousSheetMessage,
  detectSheetSignature,
  missingSheetChatMessage,
  pickBestSheet,
  resolveSheetsForRecon,
  scoreSheetRole,
} from './sheet-detector';

describe('sheet-detector (signature-based)', () => {
  it('detects purchase register via supplier GSTIN', () => {
    const r = detectSheetSignature([
      'Supplier GSTIN',
      'Invoice No',
      'Invoice Date',
      'Taxable Value',
      'CGST',
      'SGST',
    ]);
    expect(r.signature).toBe('purchase_register');
    expect(r.confidence).toBeGreaterThanOrEqual(0.7);
  });

  it('detects sales register via recipient GSTIN', () => {
    const r = detectSheetSignature([
      'Recipient GSTIN',
      'Invoice Number',
      'Invoice Date',
      'Taxable Amount',
      'Supply Category',
    ]);
    expect(r.signature).toBe('sales_register');
    expect(r.confidence).toBeGreaterThanOrEqual(0.7);
  });

  it('does not misfire sales register with stray supplier gstin as purchase when recipient present', () => {
    const r = detectSheetSignature([
      'Recipient GSTIN',
      'Invoice No',
      'Invoice Date',
      'Taxable Value',
      'Supplier GSTIN',
      'Supply Category',
    ]);
    expect(r.signature).toBe('sales_register');
  });

  it('detects GSTR-2B', () => {
    const r = detectSheetSignature([
      'GSTIN of supplier',
      'Invoice number',
      'Invoice Date',
      'Document Type',
      'ITC Available',
    ]);
    expect(r.signature).toBe('gstr_2b');
  });

  it('detects GSTR-1', () => {
    const r = detectSheetSignature([
      'GSTIN/UIN of Recipient',
      'Invoice Number',
      'Invoice Date',
      'IRN',
      'Taxable Value',
      'Supply Category',
    ]);
    expect(r.signature).toBe('gstr_1');
  });

  it('resolveSheetsForRecon classifies not_found / resolved / ambiguous', () => {
    const sheets = [
      {
        sheetName: 'Purchases Apr',
        headers: ['Supplier GSTIN', 'Invoice No', 'Invoice Date', 'Taxable Value'],
      },
      {
        sheetName: 'Purchases Mar',
        headers: ['Supplier GSTIN', 'Invoice No', 'Invoice Date', 'Taxable Value'],
      },
      {
        sheetName: 'GSTR-2B',
        headers: [
          'GSTIN of supplier',
          'Invoice number',
          'Invoice Date',
          'Document Type',
          'ITC Available',
        ],
      },
    ];
    const resolved = resolveSheetsForRecon(sheets, ['purchase_register', 'gstr_2b']);
    expect(resolved.gstr_2b.status).toBe('resolved');
    expect(resolved.purchase_register.status).toBe('ambiguous');
    expect(resolved.purchase_register.found.length).toBeGreaterThanOrEqual(2);
  });

  it('scores GSTR-2B portal headers highly (legacy API)', () => {
    const score = scoreSheetRole(
      'Portal Export',
      ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'Document Type'],
      'GSTR2B',
    );
    expect(score).toBeGreaterThanOrEqual(3);
  });

  it('picks unique best sheet (legacy API)', () => {
    const result = pickBestSheet(
      [
        {
          sheetName: 'GSTR-2B Apr',
          headers: ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value'],
        },
        { sheetName: 'Notes', headers: ['Comment'] },
      ],
      'GSTR2B',
    );
    expect(result.best?.sheetName).toBe('GSTR-2B Apr');
  });

  it('builds missing and ambiguous messages', () => {
    expect(missingSheetChatMessage(['GSTR2B'])).toMatch(/GSTR-2B/);
    expect(
      ambiguousSheetMessage('purchase_register', [
        {
          sheetName: 'A',
          headers: [],
          headerRowIndex: 1,
          confidence: 0.9,
        },
        {
          sheetName: 'B',
          headers: [],
          headerRowIndex: 1,
          confidence: 0.9,
        },
      ]),
    ).toMatch(/more than one possible purchase register/);
  });
});
