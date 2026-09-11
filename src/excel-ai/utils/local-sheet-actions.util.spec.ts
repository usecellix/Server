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

  // TASKS.md #208 — this lane runs before tiering on every write-route turn,
  // so a sheet mentioned only as a location must not become DELETE_SHEET.
  describe('delete requests that only mention a sheet as a location (#208)', () => {
    const guideContext: WorkbookContext = {
      activeSheet: 'Purchase Register',
      sheets: [
        { sheetName: 'Purchase Register', ...emptySheet },
        { sheetName: 'Summary', ...emptySheet },
        { sheetName: 'Rows Data', ...emptySheet },
      ],
    };

    it.each([
      'Delete blank rows in the Purchase Register sheet',
      'Delete column C from this sheet',
      'Remove duplicates from the Summary sheet',
      'Remove the Narration column from the Summary sheet',
      'Delete all rows where column A is blank in this sheet',
      'Delete everything in the Summary sheet',
      'Remove the header from this sheet',
      'Delete the formatting on the Summary tab',
    ])('does not propose DELETE_SHEET for %j', (message) => {
      expect(tryLocalDeleteSheetActions(message, guideContext)).toBeNull();
    });

    it.each([
      ['Delete the Summary sheet', ['Summary']],
      ['Delete sheet Summary', ['Summary']],
      ['Remove this sheet', ['Purchase Register']],
      ['remove the tab called Summary', ['Summary']],
      ['Delete the Rows Data sheet', ['Rows Data']],
      ['Delete all the sheets except Summary', ['Purchase Register', 'Rows Data']],
    ])('still deletes the sheet for %j', (message, expected) => {
      expect(tryLocalDeleteSheetActions(message, guideContext)).toEqual(
        expected.map((sheetName) => ({ type: 'DELETE_SHEET', sheetName })),
      );
    });
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
