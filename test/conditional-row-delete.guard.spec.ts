import { guardConditionalRowDeletes } from '../src/excel-ai/utils/conditional-row-delete.guard';
import { buildFixture } from '../eval/usecase-fixture';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';
import { WorkbookContext } from '../src/types/cellix.types';

/**
 * TASKS.md #238 — reproduces the two real shapes this produced in live runs
 * against a sheet with NO blank rows: `DELETE_ROW row:10 rowCount:21` (seen in
 * the task pane) and `DELETE_ROW row:1 rowCount:30` (seen in the eval harness).
 * Both were reported "verified: true" by the agentic loop.
 */
describe('conditional row delete guard (#234)', () => {
  const context = buildFixture().workbookContext as unknown as WorkbookContext;
  const del = (row: number, rowCount: number): SheetActionPayload =>
    ({ type: 'DELETE_ROW', sheetName: 'Purchase Register', row, rowCount }) as SheetActionPayload;

  it('blocks the 21-row delete seen in the task pane', () => {
    const result = guardConditionalRowDeletes(
      [del(10, 21)],
      'Delete blank rows in the Purchase Register sheet',
      context,
    );
    expect(result.actions).toEqual([]);
    expect(result.dropped[0]).toMatch(/non-empty/);
  });

  it('blocks the delete-everything shape seen in the harness', () => {
    const result = guardConditionalRowDeletes([del(1, 30)], 'Delete blank rows', context);
    expect(result.actions).toEqual([]);
  });

  it('blocks a non-blankness condition outright and points at DELETE_MATCHING_ROWS', () => {
    const result = guardConditionalRowDeletes(
      [del(4, 2)],
      'Delete rows where the supplier is ABC Traders',
      context,
    );
    expect(result.actions).toEqual([]);
    expect(result.dropped[0]).toMatch(/DELETE_MATCHING_ROWS/);
  });

  it('fails closed when the sheet data cannot be read', () => {
    const result = guardConditionalRowDeletes([del(4, 2)], 'delete the blank rows', undefined);
    expect(result.actions).toEqual([]);
  });

  // The fixture's blank GSTINs are data rows 4, 13 and 22 (Excel rows 6, 15, 24)
  // — blank in column D only, not whole-row blank. The first version of this
  // guard blocked these, turning a correct delete into "Something went wrong".
  it('allows a column-scoped blank delete that really does target blank cells', () => {
    const result = guardConditionalRowDeletes(
      [del(4, 1), del(13, 1), del(22, 1)],
      'Delete all rows where column D is blank',
      context,
    );
    expect(result.actions).toHaveLength(3);
    expect(result.dropped).toEqual([]);
  });

  it('still blocks a column-scoped delete aimed at populated cells', () => {
    const result = guardConditionalRowDeletes(
      [del(2, 1)],
      'Delete all rows where column D is blank',
      context,
    );
    expect(result.actions).toEqual([]);
  });

  it('resolves the column by header name too', () => {
    const result = guardConditionalRowDeletes(
      [del(4, 1)],
      'Delete the rows where the GSTIN is blank',
      context,
    );
    expect(result.actions).toHaveLength(1);
  });

  it('leaves an explicitly numbered delete alone', () => {
    const actions = [del(6, 1)];
    expect(guardConditionalRowDeletes(actions, 'Delete row 7', context).actions).toEqual(actions);
    expect(guardConditionalRowDeletes(actions, 'Delete rows 10-12', context).actions).toEqual(actions);
  });

  it('leaves unrelated actions and non-conditional requests alone', () => {
    const format = { type: 'FORMAT_RANGE', sheetName: 'S', row: 0, col: 0 } as SheetActionPayload;
    expect(
      guardConditionalRowDeletes([format, del(3, 1)], 'Make the header bold', context).actions,
    ).toHaveLength(2);
  });

  it('allows a blank-row delete that really does target blank rows', () => {
    const blankRowContext = {
      activeSheet: 'S',
      sheets: [
        {
          sheetName: 'S',
          usedRange: 'A1:B4',
          rowCount: 4,
          colCount: 2,
          headers: ['A', 'B'],
          sampleData: [
            ['x', 1],
            ['', ''],
            ['y', 2],
          ],
        },
      ],
    } as unknown as WorkbookContext;

    const result = guardConditionalRowDeletes(
      [{ type: 'DELETE_ROW', sheetName: 'S', row: 2, rowCount: 1 } as SheetActionPayload],
      'delete blank rows',
      blankRowContext,
    );
    expect(result.actions).toHaveLength(1);
  });
});
