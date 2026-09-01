import { WorkbookContext } from '../src/agents/types/agent.types';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import {
  beforeStateToInverseActions,
  captureStructuralOps,
  excludeStructurallyOwnedChanges,
  generateDiff,
  snapshotBeforeState,
  structuralOpsToInverseActions,
} from '../src/audit/diff.engine';

const baseContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:C2',
      rowCount: 2,
      columnCount: 3,
      values: [
        ['Item', 'Qty', 'Price'],
        ['Apple', 10, 1.5],
      ],
      formulas: [['', '', ''], ['', '', '']],
      numberFormats: [['General', 'General', 'General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

/** Full revert pipeline, mirroring change-set.service.ts's revert(): structural pre, cell-level, structural post. */
function revertRoundTrip(
  before: ReturnType<typeof buildShadowWorkbook>,
  actions: unknown[],
) {
  const beforeState = snapshotBeforeState(before);
  const after = virtualApply(before, actions as never);
  const structuralOps = captureStructuralOps(before, after, actions as never);
  const rawChanges = generateDiff(before, after);
  const changes = excludeStructurallyOwnedChanges(rawChanges, structuralOps);
  const cellInverse = beforeStateToInverseActions(beforeState, changes);
  const { pre, post } = structuralOpsToInverseActions(structuralOps);
  const inverse = [...pre, ...cellInverse, ...post];
  const restored = virtualApply(after, inverse as never);
  return { after, restored, structuralOps, inverse };
}

describe('diff.engine — structural inverse: ADD_SHEET / DELETE_SHEET (TASKS.md #12)', () => {
  it('reverting an ADD_SHEET removes the created sheet and leaves the original sheet untouched', () => {
    const before = buildShadowWorkbook(baseContext);
    const { after, restored, structuralOps } = revertRoundTrip(before, [
      { type: 'ADD_SHEET', name: 'February' },
    ]);

    expect(after.sheets.has('February')).toBe(true);
    expect(structuralOps).toEqual([
      expect.objectContaining({ opType: 'ADD_SHEET', sheetName: 'February' }),
    ]);

    expect(restored.sheets.has('February')).toBe(false);
    expect(restored.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item');
    expect(restored.sheets.get('Sheet1')?.cells.get('B2')?.value).toBe(10);
  });

  it('reverting a DELETE_SHEET recreates the sheet with its original cell values, formulas, and formats', () => {
    const before = buildShadowWorkbook(baseContext);
    // Give Sheet1 a formula and a non-default format so the restore has to carry more than plain values.
    const withFormula = virtualApply(before, [
      {
        type: 'SET_CELL',
        sheetName: 'Sheet1',
        row: 1,
        col: 2,
        value: 0.15,
        format: { numberFormat: '0.00%' },
      },
    ] as never);

    const { after, restored } = revertRoundTrip(withFormula, [
      { type: 'DELETE_SHEET', sheetName: 'Sheet1' },
    ]);

    expect(after.sheets.has('Sheet1')).toBe(false);

    const restoredSheet = restored.sheets.get('Sheet1');
    expect(restoredSheet).toBeDefined();
    expect(restoredSheet?.cells.get('A1')?.value).toBe('Item');
    expect(restoredSheet?.cells.get('C2')?.value).toBe(0.15);
    expect(restoredSheet?.cells.get('C2')?.numberFormat).toBe('0.00%');
  });

});

describe('diff.engine — structural inverse: INSERT_COLUMN / DELETE_COLUMN (TASKS.md #13)', () => {
  it('reverting an INSERT_COLUMN removes the inserted column and leaves original columns/data untouched (spec-14 repro)', () => {
    const before = buildShadowWorkbook(baseContext);
    const { after, restored, structuralOps } = revertRoundTrip(before, [
      { type: 'INSERT_COLUMN', sheetName: 'Sheet1', beforeColumn: 'B', count: 1 },
    ]);

    // Forward: B/C shifted right to C/D, new empty column at B.
    expect(after.sheets.get('Sheet1')?.columnCount).toBe(4);
    expect(after.sheets.get('Sheet1')?.cells.get('C1')?.value).toBe('Qty');
    expect(structuralOps).toEqual([
      expect.objectContaining({ opType: 'INSERT_COLUMN', sheetName: 'Sheet1' }),
    ]);

    const restoredSheet = restored.sheets.get('Sheet1');
    expect(restoredSheet?.columnCount).toBe(3);
    expect(restoredSheet?.cells.get('A1')?.value).toBe('Item');
    expect(restoredSheet?.cells.get('B1')?.value).toBe('Qty');
    expect(restoredSheet?.cells.get('C1')?.value).toBe('Price');
    expect(restoredSheet?.cells.get('B2')?.value).toBe(10);
    expect(restoredSheet?.cells.get('C2')?.value).toBe(1.5);
  });

  it('reverting a DELETE_COLUMN restores the deleted column with its original values and shifts others back', () => {
    const before = buildShadowWorkbook(baseContext);
    const { after, restored } = revertRoundTrip(before, [
      { type: 'DELETE_COLUMN', sheetName: 'Sheet1', columns: ['B'] },
    ]);

    // Forward: B (Qty) removed, C (Price) shifted left to B.
    expect(after.sheets.get('Sheet1')?.columnCount).toBe(2);
    expect(after.sheets.get('Sheet1')?.cells.get('B1')?.value).toBe('Price');

    const restoredSheet = restored.sheets.get('Sheet1');
    expect(restoredSheet?.columnCount).toBe(3);
    expect(restoredSheet?.cells.get('A1')?.value).toBe('Item');
    expect(restoredSheet?.cells.get('B1')?.value).toBe('Qty');
    expect(restoredSheet?.cells.get('B2')?.value).toBe(10);
    expect(restoredSheet?.cells.get('C1')?.value).toBe('Price');
    expect(restoredSheet?.cells.get('C2')?.value).toBe(1.5);
  });
});

const rowContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:B4',
      rowCount: 4,
      columnCount: 2,
      values: [
        ['Item', 'Qty'],
        ['Apple', 10],
        ['Banana', 20],
        ['Cherry', 30],
      ],
      formulas: [['', ''], ['', ''], ['', ''], ['', '']],
      numberFormats: [['General', 'General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('diff.engine — structural inverse: INSERT_ROW / DELETE_ROW (TASKS.md #14)', () => {
  it('reverting an INSERT_ROW removes the inserted row and leaves original rows/data untouched', () => {
    const before = buildShadowWorkbook(rowContext);
    const { after, restored, structuralOps } = revertRoundTrip(before, [
      { type: 'INSERT_ROW', sheetName: 'Sheet1', row: 2 }, // insert after 0-based row 2 (Banana) -> new row at Excel row 3
    ]);

    // Forward: Cherry (was row 4) shifts to row 5, new blank row at 3.
    expect(after.sheets.get('Sheet1')?.rowCount).toBe(5);
    expect(after.sheets.get('Sheet1')?.cells.get('A5')?.value).toBe('Cherry');
    expect(structuralOps).toEqual([
      expect.objectContaining({ opType: 'INSERT_ROW', sheetName: 'Sheet1' }),
    ]);

    const restoredSheet = restored.sheets.get('Sheet1');
    expect(restoredSheet?.rowCount).toBe(4);
    expect(restoredSheet?.cells.get('A1')?.value).toBe('Item');
    expect(restoredSheet?.cells.get('A2')?.value).toBe('Apple');
    expect(restoredSheet?.cells.get('A3')?.value).toBe('Banana');
    expect(restoredSheet?.cells.get('A4')?.value).toBe('Cherry');
    expect(restoredSheet?.cells.get('B4')?.value).toBe(30);
  });

  it('reverting a DELETE_ROW restores the deleted row with its original values and shifts others back', () => {
    const before = buildShadowWorkbook(rowContext);
    const { after, restored } = revertRoundTrip(before, [
      { type: 'DELETE_ROW', sheetName: 'Sheet1', rows: [3] }, // 1-based: Banana
    ]);

    // Forward: Banana removed, Cherry shifts from row 4 to row 3.
    expect(after.sheets.get('Sheet1')?.rowCount).toBe(3);
    expect(after.sheets.get('Sheet1')?.cells.get('A3')?.value).toBe('Cherry');

    const restoredSheet = restored.sheets.get('Sheet1');
    expect(restoredSheet?.rowCount).toBe(4);
    expect(restoredSheet?.cells.get('A1')?.value).toBe('Item');
    expect(restoredSheet?.cells.get('A2')?.value).toBe('Apple');
    expect(restoredSheet?.cells.get('A3')?.value).toBe('Banana');
    expect(restoredSheet?.cells.get('B3')?.value).toBe(20);
    expect(restoredSheet?.cells.get('A4')?.value).toBe('Cherry');
    expect(restoredSheet?.cells.get('B4')?.value).toBe(30);
  });

  it('reverting a multi-row DELETE_ROW restores all deleted rows in the correct order', () => {
    const before = buildShadowWorkbook(rowContext);
    const { restored } = revertRoundTrip(before, [
      { type: 'DELETE_ROW', sheetName: 'Sheet1', rows: [2, 4] }, // Apple and Cherry
    ]);

    const restoredSheet = restored.sheets.get('Sheet1');
    expect(restoredSheet?.rowCount).toBe(4);
    expect(restoredSheet?.cells.get('A1')?.value).toBe('Item');
    expect(restoredSheet?.cells.get('A2')?.value).toBe('Apple');
    expect(restoredSheet?.cells.get('A3')?.value).toBe('Banana');
    expect(restoredSheet?.cells.get('A4')?.value).toBe('Cherry');
  });
});

describe('diff.engine — structural inverse: CREATE_TABLE (TASKS.md #16)', () => {
  it('reverting a CREATE_TABLE produces a DELETE_TABLE inverse targeting the same sheet/table, with cell values untouched', () => {
    const before = buildShadowWorkbook(baseContext);
    const { after, restored, structuralOps, inverse } = revertRoundTrip(before, [
      { type: 'CREATE_TABLE', sheetName: 'Sheet1', range: 'A1:C2', tableName: 'ItemsTable', hasHeaders: true },
    ]);

    // CREATE_TABLE wraps an existing range — it does not touch cell values, so the
    // shadow (which has no table-state concept, see virtual-apply-catalog.ts) is
    // byte-identical before/after. The real assertion here is the inverse *action
    // shape* that gets returned to the frontend for a real Office.js convertToRange().
    expect(structuralOps).toEqual([
      expect.objectContaining({
        opType: 'CREATE_TABLE',
        sheetName: 'Sheet1',
        params: { tableName: 'ItemsTable' },
      }),
    ]);
    expect(inverse).toContainEqual({
      type: 'DELETE_TABLE',
      sheetName: 'Sheet1',
      tableName: 'ItemsTable',
    });

    expect(restored.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item');
    expect(restored.sheets.get('Sheet1')?.cells.get('B2')?.value).toBe(10);
    expect(after.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item');
  });

  it('produces no structural op when tableName or sheetName is missing', () => {
    const before = buildShadowWorkbook(baseContext);
    const after = virtualApply(before, [
      { type: 'CREATE_TABLE', sheetName: 'Sheet1', range: 'A1:C2', tableName: '', hasHeaders: true },
    ] as never);
    const ops = captureStructuralOps(before, after, [
      { type: 'CREATE_TABLE', sheetName: 'Sheet1', range: 'A1:C2', tableName: '', hasHeaders: true },
    ] as never);
    expect(ops).toEqual([]);
  });
});

describe('diff.engine — structural inverse: CREATE_CHART (TASKS.md #15)', () => {
  it('captures a pending structuralOp (no chartId yet) for a chart create', () => {
    const before = buildShadowWorkbook(baseContext);
    const actions = [
      {
        type: 'CREATE_CHART',
        sheetName: 'Sheet1',
        sourceSheetName: 'Sheet1',
        sourceRange: 'A1:B2',
        chartType: 'column',
      },
    ];
    const after = virtualApply(before, actions as never);
    const ops = captureStructuralOps(before, after, actions as never);

    expect(ops).toEqual([
      expect.objectContaining({
        opType: 'CREATE_CHART',
        sheetName: 'Sheet1',
        params: { sourceRange: 'A1:B2' },
      }),
    ]);
    expect(ops[0]?.params.chartId).toBeUndefined();
  });

  it('throws when building the inverse if chartId was never patched in (fails closed, TASKS.md #15)', () => {
    const before = buildShadowWorkbook(baseContext);
    const actions = [
      {
        type: 'CREATE_CHART',
        sheetName: 'Sheet1',
        sourceSheetName: 'Sheet1',
        sourceRange: 'A1:B2',
        chartType: 'column',
      },
    ];
    const after = virtualApply(before, actions as never);
    const ops = captureStructuralOps(before, after, actions as never);
    expect(() => structuralOpsToInverseActions(ops)).toThrow(/no chartId was ever reported/);
  });

  it('produces a DELETE_CHART inverse once chartId is patched in — matching markApplied()\'s own patch shape', () => {
    const before = buildShadowWorkbook(baseContext);
    const actions = [
      {
        type: 'CREATE_CHART',
        sheetName: 'Sheet1',
        sourceSheetName: 'Sheet1',
        sourceRange: 'A1:B2',
        chartType: 'column',
      },
    ];
    const after = virtualApply(before, actions as never);
    const ops = captureStructuralOps(before, after, actions as never);
    const patched = ops.map((op) => ({ ...op, params: { ...op.params, chartId: 'Chart_real_1' } }));

    const { pre, post } = structuralOpsToInverseActions(patched);
    expect(pre).toEqual([]);
    expect(post).toContainEqual({
      type: 'DELETE_CHART',
      sheetName: 'Sheet1',
      chartId: 'Chart_real_1',
    });
  });
});

describe('diff.engine — structural inverse: CONDITIONAL_FORMAT (TASKS.md #40)', () => {
  it('captures a pending structuralOp (no ruleId yet) for a plain create', () => {
    const before = buildShadowWorkbook(baseContext);
    const actions = [
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        range: 'B2:B2',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 5, format: { fillColor: '#FFC7CE' } },
      },
    ];
    const after = virtualApply(before, actions as never);
    const ops = captureStructuralOps(before, after, actions as never);

    expect(ops).toEqual([
      expect.objectContaining({
        opType: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        params: { range: 'B2:B2' },
      }),
    ]);
    expect(ops[0]?.params.ruleId).toBeUndefined();
  });

  it('captures nothing for a MODIFY (existingRuleId set) — only a create needs a revert entry', () => {
    const before = buildShadowWorkbook(baseContext);
    const actions = [
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        range: 'B2:B2',
        existingRuleId: 'cf-existing',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 5, format: { fillColor: '#FFC7CE' } },
      },
    ];
    const after = virtualApply(before, actions as never);
    const ops = captureStructuralOps(before, after, actions as never);
    expect(ops).toEqual([]);
  });

  it('throws when building the inverse if ruleId was never patched in (fails closed, TASKS.md #40)', () => {
    const before = buildShadowWorkbook(baseContext);
    const actions = [
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        range: 'B2:B2',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 5, format: { fillColor: '#FFC7CE' } },
      },
    ];
    const after = virtualApply(before, actions as never);
    const ops = captureStructuralOps(before, after, actions as never);
    expect(() => structuralOpsToInverseActions(ops)).toThrow(/no ruleId was ever reported/);
  });

  it('produces a DELETE_CONDITIONAL_FORMAT inverse once ruleId is patched in — matching markApplied()\'s own patch shape', () => {
    const before = buildShadowWorkbook(baseContext);
    const actions = [
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        range: 'B2:B2',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 5, format: { fillColor: '#FFC7CE' } },
      },
    ];
    const after = virtualApply(before, actions as never);
    const ops = captureStructuralOps(before, after, actions as never);
    const patched = ops.map((op) => ({ ...op, params: { ...op.params, ruleId: 'cf-real-1' } }));

    const { pre, post } = structuralOpsToInverseActions(patched);
    expect(pre).toEqual([]);
    expect(post).toContainEqual({
      type: 'DELETE_CONDITIONAL_FORMAT',
      sheetName: 'Sheet1',
      ruleId: 'cf-real-1',
    });
  });
});

describe('diff.engine — structural inverse: MERGE_CELLS / UNMERGE_CELLS (TASKS.md #17)', () => {
  it('reverting a MERGE_CELLS unmerges the range and restores the discarded (non-top-left) values', () => {
    const before = buildShadowWorkbook(baseContext);
    const { after, restored, structuralOps, inverse } = revertRoundTrip(before, [
      { type: 'MERGE_CELLS', sheetName: 'Sheet1', range: 'A1:B1' },
    ]);

    // Forward: top-left (A1) survives, B1's value is discarded (real Excel merge behavior, TASKS.md #66).
    expect(after.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item');
    expect(after.sheets.get('Sheet1')?.cells.get('B1')?.value).toBeNull();
    expect(structuralOps).toEqual([
      expect.objectContaining({ opType: 'MERGE_CELLS', sheetName: 'Sheet1', params: { range: 'A1:B1' } }),
    ]);
    expect(inverse[0]).toEqual({ type: 'UNMERGE_CELLS', sheetName: 'Sheet1', range: 'A1:B1' });

    const restoredSheet = restored.sheets.get('Sheet1');
    expect(restoredSheet?.cells.get('A1')?.value).toBe('Item');
    expect(restoredSheet?.cells.get('B1')?.value).toBe('Qty');
  });

  it('reverting an UNMERGE_CELLS produces a MERGE_CELLS inverse targeting the same range', () => {
    const before = buildShadowWorkbook(baseContext);
    const { structuralOps, inverse } = revertRoundTrip(before, [
      { type: 'UNMERGE_CELLS', sheetName: 'Sheet1', range: 'A1:B1' },
    ]);

    expect(structuralOps).toEqual([
      expect.objectContaining({ opType: 'UNMERGE_CELLS', sheetName: 'Sheet1', params: { range: 'A1:B1' } }),
    ]);
    expect(inverse).toContainEqual({ type: 'MERGE_CELLS', sheetName: 'Sheet1', range: 'A1:B1' });
  });
});

describe('diff.engine — structural op edge case', () => {
  it('produces no structural ops for a batch with no structural actions', () => {
    const before = buildShadowWorkbook(baseContext);
    const actions = [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 0, col: 0, value: 'Product' },
    ];
    const after = virtualApply(before, actions as never);
    const ops = captureStructuralOps(before, after, actions as never);
    expect(ops).toEqual([]);
  });
});
