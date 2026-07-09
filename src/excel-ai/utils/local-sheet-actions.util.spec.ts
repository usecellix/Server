import {
  buildDeleteSheetAnswer,
  extractDeleteSheetNames,
  tryLocalDeleteSheetActions,
} from './local-sheet-actions.util';
import { WorkbookContext } from '../../types/cellix.types';

const context: WorkbookContext = {
  activeSheet: 'Invoices',
  sheets: [
    { sheetName: 'Invoices', usedRange: 'A1', headers: [], sampleData: [], rowCount: 1, colCount: 1, columnMeta: [] },
    { sheetName: 'Azhar', usedRange: 'A1', headers: [], sampleData: [], rowCount: 1, colCount: 1, columnMeta: [] },
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
});
