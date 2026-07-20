import { validateCrossSheetActions } from '../src/sheets/multi-sheet.service';
import { SheetSnapshot, WorkbookContext } from '../src/types/cellix.types';

function sheet(name: string): SheetSnapshot {
  return {
    sheetName: name,
    usedRange: 'A1',
    rowCount: 1,
    colCount: 1,
    headers: [],
    sampleData: [],
  };
}

describe('validateCrossSheetActions', () => {
  const context: WorkbookContext = {
    activeSheet: 'Sheet1',
    sheets: [sheet('Sheet1'), sheet('Sheet2')],
  };

  it('keeps ADD_ROW for a sheet created in the same batch', () => {
    const result = validateCrossSheetActions(
      [
        { type: 'CREATE_SHEET', name: 'Tax Returns' },
        { type: 'ADD_ROW', sheetName: 'Tax Returns', data: ['A', 1] },
        { type: 'ADD_ROW', sheetName: 'Tax Returns', data: ['B', 2] },
        { type: 'SET_FORMULA', sheetName: 'Tax Returns', row: 1, col: 2, formula: '=A2+B2' },
      ],
      context,
    );

    expect(result.errors).toEqual([]);
    expect(result.invalid).toHaveLength(0);
    expect(result.valid).toHaveLength(4);
  });

  it('rejects actions for sheets that are neither existing nor created', () => {
    const result = validateCrossSheetActions(
      [{ type: 'ADD_ROW', sheetName: 'Missing', data: [1] }],
      context,
    );

    expect(result.valid).toHaveLength(0);
    expect(result.errors[0]).toContain('Missing');
  });

  it('accepts ADD_SHEET + follow-up using name field', () => {
    const result = validateCrossSheetActions(
      [
        { type: 'ADD_SHEET', name: 'Tax Returns' },
        { type: 'SET_CELL', sheetName: 'Tax Returns', row: 0, col: 0, value: 'Header' },
      ],
      context,
    );

    expect(result.errors).toEqual([]);
    expect(result.valid).toHaveLength(2);
  });
});
