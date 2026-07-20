import { buildShadowWorkbook, shadowSheetToContext } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import { WorkbookContext } from '../src/agents/types/agent.types';

describe('virtualApply AGGREGATE_TABLE', () => {
  it('writes top-N supplier totals to the destination sheet', () => {
    const context: WorkbookContext = {
      activeSheetName: 'Purchase Register',
      sheets: [
        {
          name: 'Purchase Register',
          usedRange: 'A1:B5',
          rowCount: 5,
          columnCount: 2,
          values: [
            ['Supplier', 'Total Amount'],
            ['Acme', 100],
            ['Beta', 50],
            ['Acme', 25],
            ['Gamma', 200],
          ],
          formulas: [[], [], [], [], []],
          numberFormats: [[], [], [], [], []],
          structure: 'data_table',
        },
      ],
      namedRanges: [],
      tables: [],
    };

    const after = virtualApply(buildShadowWorkbook(context), [
      {
        type: 'AGGREGATE_TABLE',
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:B5',
        groupByColumn: 'Supplier',
        aggregations: [{ column: 'Total Amount', fn: 'sum', outputLabel: 'Total Spend' }],
        sortBy: { column: 'Total Spend', direction: 'desc' },
        topN: 2,
        destSheet: 'Dashboard',
        destStartCell: 'A4',
        hasHeaders: true,
      },
    ]);

    const dest = shadowSheetToContext(after.sheets.get('Dashboard')!);
    // Rows are 0-indexed in snapshot; A4 is row index 3
    expect(dest.values[3]).toEqual(['Supplier', 'Total Spend']);
    expect(dest.values[4]).toEqual(['Gamma', 200]);
    expect(dest.values[5]).toEqual(['Acme', 125]);
  });
});
