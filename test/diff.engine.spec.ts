import { WorkbookContext } from '../src/agents/types/agent.types';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import {
  beforeStateToInverseActions,
  generateDiff,
  snapshotBeforeState,
} from '../src/audit/diff.engine';
import { CellChange } from '../src/audit/types/change-set.types';
import { Action } from '../src/agents/types/agent.types';

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
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('diff.engine', () => {
  it('snapshots before state from shadow workbook', () => {
    const shadow = buildShadowWorkbook(baseContext);
    const state = snapshotBeforeState(shadow);
    expect(state['Sheet1!A1']).toEqual({ value: 'Item', formula: '', format: 'General' });
    expect(state['Sheet1!C2']).toEqual({ value: 1.5, formula: '', format: 'General' });
  });

  it('generates cell-level diff after virtual apply', () => {
    const before = buildShadowWorkbook(baseContext);
    const beforeState = snapshotBeforeState(before);
    const actions = [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: 12 },
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=B2*1.1' },
    ];
    const after = virtualApply(before, actions as never);
    const changes = generateDiff(before, after);

    expect(changes.length).toBeGreaterThanOrEqual(2);
    const qtyChange = changes.find((c: CellChange) => c.cell === 'B2');
    expect(qtyChange).toMatchObject({ sheet: 'Sheet1', before: 10, after: 12 });

    const inverse = beforeStateToInverseActions(beforeState, changes);
    expect(inverse.some((a: Action) => a.type === 'SET_CELL' && (a as Action & { address?: string }).address === 'B2')).toBe(true);
  });

  it('produces inverse actions that restore before values', () => {
    const before = buildShadowWorkbook(baseContext);
    const beforeState = snapshotBeforeState(before);
    const after = virtualApply(before, [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 0, col: 0, value: 'Product' },
    ] as never);
    const changes = generateDiff(before, after);
    const inverse = beforeStateToInverseActions(beforeState, changes);

    const restored = virtualApply(after, inverse as never);
    expect(restored.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item');
  });
});
