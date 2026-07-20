import { Tier0DirectService } from '../src/excel-ai/services/tier0-direct.service';
import { WorkbookContext } from '../src/agents/types/agent.types';

const workbookContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:D10',
      rowCount: 10,
      columnCount: 4,
      values: [['Name', 'Amount'], ['A', 100]],
      formulas: [],
      numberFormats: [],
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('Tier0DirectService', () => {
  const service = new Tier0DirectService();

  it('resolves CELL_FORMAT with explicit cell reference', () => {
    const result = service.resolve('CELL_FORMAT', 'bold A1:C1', workbookContext);
    expect(result).toEqual({
      skippedLLM: true,
      actions: [
        {
          type: 'FORMAT_RANGE',
          sheetName: 'Sheet1',
          row: 0,
          col: 0,
          rowCount: 1,
          colCount: 3,
          format: { bold: true },
        },
      ],
    });
  });

  it('resolves FREEZE_PANES without LLM', () => {
    const result = service.resolve('FREEZE_PANES', 'freeze top row', workbookContext);
    expect(result?.actions[0]).toMatchObject({
      type: 'FREEZE_PANES',
      freezeRows: 1,
      freezeColumns: 0,
    });
  });

  it('resolves VISIBILITY_TOGGLE for hide column', () => {
    const result = service.resolve('VISIBILITY_TOGGLE', 'hide column F', workbookContext);
    expect(result?.actions[0]).toMatchObject({
      type: 'HIDE_COLUMN',
      col: 5,
      colCount: 1,
    });
  });

  it('resolves ROW_COL_STRUCTURE for delete column', () => {
    const result = service.resolve('ROW_COL_STRUCTURE', 'delete column C', workbookContext);
    expect(result?.actions[0]).toMatchObject({
      type: 'DELETE_COLUMN',
      col: 2,
      colCount: 1,
    });
  });

  it('returns null for implicit cell-format target (no explicit A1 reference)', () => {
    const result = service.resolve('CELL_FORMAT', 'bold the header row', workbookContext);
    expect(result).toBeNull();
  });

  it('returns null for insert row without explicit row reference', () => {
    const result = service.resolve('ROW_COL_STRUCTURE', 'insert a row', workbookContext);
    expect(result).toBeNull();
  });
});
