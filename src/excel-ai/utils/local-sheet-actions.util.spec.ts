import {
  buildDeleteSheetAnswer,
  extractDeleteSheetNames,
  tryLocalDeleteSheetActions,
} from './local-sheet-actions.util';
import { WorkbookContext } from '../../types/cellix.types';

const emptySheet = {
  usedRange: 'A1',
  headers: [] as string[],
  sampleData: [] as (string | number | null)[][],
  rowCount: 1,
  colCount: 1,
  columnMeta: [] as never[],
};

const context: WorkbookContext = {
  activeSheet: 'Invoices',
  sheets: [
    { sheetName: 'Invoices', ...emptySheet },
    { sheetName: 'Azhar', ...emptySheet },
  ],
};

const multiSheetContext: WorkbookContext = {
  activeSheet: 'Purchase Register',
  sheets: [
    { sheetName: 'Purchase Register', ...emptySheet },
    { sheetName: 'Sales', ...emptySheet },
    { sheetName: 'Invoices', ...emptySheet },
  ],
};

describe('local-sheet-actions.util', () => {
  it('extracts sheet name from @[mention] tags', () => {
    expect(
      extractDeleteSheetNames('Delete the sheet Azhar @[Azhar]', ['Invoices', 'Azhar']),
    ).toEqual(['Azhar']);
  });

  it('returns deterministic delete actions', () => {
    const actions = tryLocalDeleteSheetActions('Delete the sheet Azhar @[Azhar]', context);
    expect(actions).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Azhar' }]);
  });

  it('builds delete answer text', () => {
    expect(buildDeleteSheetAnswer(['Azhar'])).toContain('Azhar');
  });

  it('deletes all sheets except the preserved one', () => {
    const actions = tryLocalDeleteSheetActions(
      'Delete all the sheets except purchase register',
      multiSheetContext,
    );
    expect(actions).toEqual([
      { type: 'DELETE_SHEET', sheetName: 'Sales' },
      { type: 'DELETE_SHEET', sheetName: 'Invoices' },
    ]);
  });

  it('does not invert for plain delete of a named sheet', () => {
    const actions = tryLocalDeleteSheetActions(
      'Delete sheet Purchase Register',
      multiSheetContext,
    );
    expect(actions).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Purchase Register' }]);
  });

  it('returns null when only the keep sheet exists (nothing to delete)', () => {
    const onlyKeep: WorkbookContext = {
      activeSheet: 'Purchase Register',
      sheets: [{ sheetName: 'Purchase Register', ...emptySheet }],
    };
    expect(
      tryLocalDeleteSheetActions('Delete all the sheets except purchase register', onlyKeep),
    ).toBeNull();
  });
});
