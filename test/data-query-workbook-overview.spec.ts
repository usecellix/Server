import {
  buildDataQueryUserMessage,
  formatWorkbookOverview,
} from '../src/excel-ai/prompts/data-query-system-prompt';
import { WorkbookContext } from '../src/types/cellix.types';

/**
 * TASKS.md #218 — the read-only lane was handed only the active sheet's sliced
 * rows, so guide Q&A.2 questions (LIST_SHEETS, LIST_NAMED_RANGES,
 * DESCRIBE_SHEET) were answered with "I can only see the Purchase Register
 * sheet" even though the request carried all four sheets and their metadata.
 */
describe('data-query workbook overview (#218)', () => {
  const context = {
    activeSheet: 'Purchase Register',
    sheets: [
      { sheetName: 'Purchase Register', usedRange: 'A1:I31', rowCount: 31, colCount: 9, headers: ['Invoice No', 'Amount'], sampleData: [], formulaSummary: 'Total formulas: 90. Functions used: IF, ROUND, LEFT' },
      { sheetName: 'GSTR-2A', usedRange: 'A1:D3', rowCount: 3, colCount: 4, headers: ['GSTIN'], sampleData: [] },
      { sheetName: 'Summary', usedRange: 'A1:B2', rowCount: 2, colCount: 2, headers: ['Particulars'], sampleData: [] },
      { sheetName: 'Working', usedRange: 'A1:A2', rowCount: 2, colCount: 1, headers: ['Note'], sampleData: [], isHidden: true },
    ],
    namedRanges: [{ name: 'TaxableAmount', formula: "='Purchase Register'!$E$2:$E$31", type: 'Range' }],
  } as unknown as WorkbookContext;

  const slice = {
    sheetName: 'Purchase Register',
    headers: ['Invoice No', 'Taxable Amount'],
    columnLetters: ['A', 'E'],
    rows: [['INV/2024/001', 118500]],
    totalRows: 30,
  } as never;

  it('lists every sheet, including hidden ones', () => {
    const overview = formatWorkbookOverview(context);
    expect(overview).toContain('4 sheet(s)');
    for (const name of ['Purchase Register', 'GSTR-2A', 'Summary', 'Working']) {
      expect(overview).toContain(name);
    }
    expect(overview).toMatch(/"Working"\s*\[hidden\]/);
  });

  it('carries named ranges and the active sheet formula summary', () => {
    const overview = formatWorkbookOverview(context);
    expect(overview).toContain('TaxableAmount');
    expect(overview).toContain('Functions used: IF, ROUND, LEFT');
  });

  it('says so explicitly when there are no named ranges', () => {
    const overview = formatWorkbookOverview({ ...context, namedRanges: [] } as WorkbookContext);
    expect(overview).toContain('none defined');
  });

  it('puts the overview in front of the data table the model reads', () => {
    const message = buildDataQueryUserMessage('How many sheets are in this workbook?', slice, context);
    expect(message.indexOf('WORKBOOK:')).toBeLessThan(message.indexOf('DATA TABLE:'));
    expect(message).toContain('USER QUESTION: How many sheets');
  });

  it('degrades to the plain message when no context is supplied', () => {
    const message = buildDataQueryUserMessage('total?', slice, undefined);
    expect(message).not.toContain('WORKBOOK:');
    expect(message).toContain('DATA TABLE:');
  });
});
