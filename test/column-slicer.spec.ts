import {
  sliceFromRawData,
  sliceRelevantColumns,
} from '../src/excel-ai/utils/column-slicer.util';
import { WorkbookContext } from '../src/types/cellix.types';

const purchaseRegisterRows: unknown[][] = [
  ['Purchase Register', '', '', '', '', '', ''],
  ['Date', 'Voucher No', 'Party Name', 'Taxable Amount', 'CGST', 'SGST', 'Total'],
  ['01-04-2024', 'INV-001', 'ABC Traders', '10000.00', '1868.41 Dr', '1868.41 Dr', '13736.82'],
  ['02-04-2024', 'INV-002', 'XYZ Pvt Ltd', '5000.00', '945.00 Dr', '945.00 Dr', '6890.00'],
  ['03-04-2024', 'INV-003', 'PQR Corp', '20000.00', '3600.00 Dr', '3600.00 Dr', '27200.00'],
  ['04-04-2024', 'INV-004', 'LMN Stores', '8000.00', '1440.00 Dr', '1440.00 Dr', '10880.00'],
];

const purchaseRegisterContext: WorkbookContext = {
  activeSheet: 'Purchase register',
  sheets: [
    {
      sheetName: 'Purchase register',
      usedRange: 'A1:G6',
      rowCount: 6,
      colCount: 7,
      headers: ['Date', 'Voucher No', 'Party Name', 'Taxable Amount', 'CGST', 'SGST', 'Total'],
      sampleData: purchaseRegisterRows.slice(1, 4) as (string | number | null)[][],
    },
  ],
};

describe('sliceRelevantColumns — keyword matching', () => {
  it('extracts CGST column for "total CGST" query', () => {
    const result = sliceFromRawData(
      'What is the total CGST in this sheet?',
      purchaseRegisterRows,
      'Purchase register',
    );

    expect(result.sheets).toHaveLength(1);
    const sheet = result.sheets[0];
    expect(sheet.headers).toContain('CGST');
    expect(sheet.headers).not.toContain('SGST');
    expect(sheet.headers).not.toContain('Taxable Amount');
  });

  it('includes anchor columns (Date, Voucher No) alongside CGST', () => {
    const result = sliceFromRawData('What is the total CGST?', purchaseRegisterRows, 'Purchase register');
    const sheet = result.sheets[0];
    expect(sheet.headers).toContain('Date');
    expect(sheet.headers).toContain('Voucher No');
  });

  it('extracts SGST column for "total SGST" query', () => {
    const result = sliceFromRawData('sum of SGST', purchaseRegisterRows, 'Purchase register');
    const sheet = result.sheets[0];
    expect(sheet.headers).toContain('SGST');
    expect(sheet.headers).not.toContain('CGST');
  });

  it('extracts amount column for "total amount" query', () => {
    const result = sliceFromRawData(
      'what is the total taxable amount',
      purchaseRegisterRows,
      'Purchase register',
    );
    const sheet = result.sheets[0];
    expect(sheet.headers).toContain('Taxable Amount');
  });

  it('returns all columns when no specific column mentioned', () => {
    const result = sliceFromRawData('show me everything', purchaseRegisterRows, 'Purchase register');
    const sheet = result.sheets[0];
    expect(sheet.headers.length).toBe(7);
  });
});

describe('sliceRelevantColumns — Tally header detection', () => {
  it('skips title row and finds real header at row 1', () => {
    const result = sliceFromRawData('total CGST', purchaseRegisterRows, 'Purchase register');
    const sheet = result.sheets[0];
    expect(sheet.rows[0][0]).toBe('01-04-2024');
    expect(sheet.totalRows).toBe(4);
  });

  it('handles sheets where row 0 is already the header', () => {
    const rows: unknown[][] = [
      ['Invoice', 'Amount', 'GST'],
      ['INV-1', '10000', '1800'],
      ['INV-2', '5000', '900'],
    ];
    const result = sliceFromRawData('total GST', rows, 'Sales');
    const sheet = result.sheets[0];
    expect(sheet.totalRows).toBe(2);
    expect(sheet.headers).toContain('GST');
  });
});

describe('sliceRelevantColumns — column letters', () => {
  it('maps column indices to correct Excel letters', () => {
    const result = sliceFromRawData('total CGST', purchaseRegisterRows, 'Purchase register');
    const sheet = result.sheets[0];
    const cgstIdx = sheet.headers.indexOf('CGST');
    expect(sheet.columnLetters[cgstIdx]).toBe('E');
  });
});

describe('sliceRelevantColumns — edge cases', () => {
  it('returns empty slice for empty workbook context', () => {
    const result = sliceRelevantColumns('total CGST', { activeSheet: 'Sheet1', sheets: [] }, []);
    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].rows).toHaveLength(0);
  });

  it('returns empty slice when sheet has no rows', () => {
    const result = sliceRelevantColumns(
      'total CGST',
      { activeSheet: 'Empty', sheets: [] },
      [],
      'Empty',
    );
    expect(result.sheets[0].rows).toHaveLength(0);
  });

  it('handles missing workbookContext gracefully', () => {
    const result = sliceRelevantColumns('total CGST', null, purchaseRegisterRows);
    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].headers).toContain('CGST');
  });

  it('uses workbook snapshot headers when provided', () => {
    const result = sliceRelevantColumns(
      'total CGST',
      purchaseRegisterContext,
      purchaseRegisterRows,
      'Purchase register',
    );
    expect(result.sheets[0].headers).toContain('CGST');
  });
});
