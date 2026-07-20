import { buildShadowWorkbook, shadowSheetToContext } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import { WorkbookContext } from '../src/agents/types/agent.types';

describe('virtualSortRange sparse / compressed sheets', () => {
  it('does not mutate a truncated sheet (null-padded rows)', () => {
    const context: WorkbookContext = {
      activeSheetName: 'Purchase Register',
      sheets: [
        {
          name: 'Purchase Register',
          usedRange: 'A1:B51',
          rowCount: 51,
          columnCount: 2,
          // Only header + 2 data rows present — middle rows are null when expanded.
          values: [
            ['Tax Amount', 'Invoice'],
            [100, 'INV-1'],
            [50, 'INV-2'],
          ],
          formulas: [[], [], []],
          numberFormats: [[], [], []],
          structure: 'data_table',
        },
      ],
      namedRanges: [],
      tables: [],
    };

    // Inflate to declared rowCount with null padding (mirrors shadowSheetToContext).
    const shadow = buildShadowWorkbook(context);
    const sheet = shadow.sheets.get('Purchase Register')!;
    sheet.rowCount = 51;

    const before = shadowSheetToContext(sheet);
    // Force sparse matrix: rowCount 51 but only first 3 rows populated in cells map.
    expect(before.values.length).toBeLessThanOrEqual(51);

    const afterShadow = virtualApply(shadow, [
      {
        type: 'SORT_RANGE',
        sheetName: 'Purchase Register',
        range: 'A1:B51',
        key: 0,
        ascending: false,
        hasHeaders: true,
      },
    ]);

    const after = shadowSheetToContext(afterShadow.sheets.get('Purchase Register')!);
    // Sparse sort must be a no-op — live Excel SORT_RANGE owns the reorder.
    expect(after.values[1]?.[0]).toBe(100);
    expect(after.values[2]?.[0]).toBe(50);
    expect(afterShadow.changedCells.size).toBe(0);
  });
});
