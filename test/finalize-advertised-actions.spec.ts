import { ConversationEngineService } from '../src/excel-ai/services/conversation-engine.service';
import { EXECUTOR_ADVERTISED_ACTION_TYPES } from '../src/excel-ai/types/action-catalog';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #215 — `sanitizeAction` ends in `default: return null`, so an action
 * type the Executor is told to use but the switch has no case for is discarded
 * AFTER the agentic loop reports "verified: true". CONDITIONAL_FORMAT,
 * DATA_VALIDATION, AUTO_FILTER and HIDE_GRIDLINES were all in that state: the
 * colour-scale and data-bar use cases failed with "Something went wrong", and
 * the dropdown use case returned an answer claiming "DATA_VALIDATION" with no
 * validation action attached.
 *
 * This pins the invariant rather than the four instances: every type offered to
 * the Executor must survive finalizeActions when given a well-formed payload.
 */
describe('every advertised action type survives finalizeActions (#215)', () => {
  // finalizeActions is pure with respect to these collaborators — none is
  // touched on this path, so the nulls keep the test to the sanitize contract.
  const engine = new ConversationEngineService(
    undefined as never,
    { hasLlmProvider: false } as never,
    undefined as never,
    undefined as never,
    undefined as never,
    undefined as never,
  );

  const analysis = {
    isEmpty: false,
    headers: ['Invoice No', 'Amount'],
    rowCount: 31,
    columnCount: 9,
  } as never;

  /** A minimally valid payload per advertised type — the shape executor.prompt.ts documents. */
  const SAMPLES: Record<string, SheetActionPayload> = {
    SET_CELL: { type: 'SET_CELL', sheetName: 'S', row: 5, col: 1, value: 'x' },
    SET_FORMULA: { type: 'SET_FORMULA', sheetName: 'S', row: 5, col: 1, formula: '=1+1' },
    HIGHLIGHT_CELL: { type: 'HIGHLIGHT_CELL', sheetName: 'S', row: 5, col: 1, color: '#FF0000' },
    BATCH_SET: { type: 'BATCH_SET', sheetName: 'S', operations: [{ address: 'B6', value: 1 }] },
    WRITE_TABLE: { type: 'WRITE_TABLE', sheetName: 'S', headers: ['a', 'b'], rows: [[1, 2]] },
    ADD_ROW: { type: 'ADD_ROW', sheetName: 'S', data: ['a', 'b'] },
    DELETE_ROW: { type: 'DELETE_ROW', sheetName: 'S', row: 5 },
    INSERT_ROW: { type: 'INSERT_ROW', sheetName: 'S', row: 5, count: 2 },
    INSERT_COLUMN: { type: 'INSERT_COLUMN', sheetName: 'S', columnName: 'New', position: 'afterLastColumn' },
    DELETE_COLUMN: { type: 'DELETE_COLUMN', sheetName: 'S', col: 3 },
    HIDE_ROW: { type: 'HIDE_ROW', sheetName: 'S', row: 4, rowCount: 2 },
    UNHIDE_ROW: { type: 'UNHIDE_ROW', sheetName: 'S', row: 4, rowCount: 2 },
    HIDE_COLUMN: { type: 'HIDE_COLUMN', sheetName: 'S', col: 8, colCount: 1 },
    UNHIDE_COLUMN: { type: 'UNHIDE_COLUMN', sheetName: 'S', col: 8, colCount: 1 },
    SET_COLUMN_WIDTH: { type: 'SET_COLUMN_WIDTH', sheetName: 'S', col: 1, width: 20 },
    FORMAT_RANGE: { type: 'FORMAT_RANGE', sheetName: 'S', row: 1, col: 0, rowCount: 30, colCount: 9, format: { bold: true } },
    // TASKS.md #258 — the shape a model actually reaches for unprompted
    // ("copy the formula in J2 down to J61" -> a single `range`), not the
    // row/col/rowCount shape the OLD sanitizeAction case (wrongly) checked
    // and this fixture used to assert against — that agreement between a
    // buggy implementation and a buggy fixture is exactly why the bug went
    // uncaught. The client's real FILL_DOWN handler needs sourceRange/
    // targetRange, never row/col at all.
    FILL_DOWN: { type: 'FILL_DOWN', sheetName: 'S', range: 'E2:E31' } as never,
    FILL_RIGHT: { type: 'FILL_RIGHT', sheetName: 'S', row: 1, col: 4, colCount: 3 },
    MERGE_CELLS: { type: 'MERGE_CELLS', sheetName: 'S', row: 1, col: 0, rowCount: 1, colCount: 3 },
    // row: 0 (not 1) is deliberate — TASKS.md #258's live repro was "unmerge
    // ALL merged cells in this sheet", which naturally spans the whole used
    // range starting at row 0. The previous row: 1 fixture never touched the
    // header-mutation guard's row===0 check at all, so it could not have
    // caught the bug that guard actually had (UNMERGE_CELLS missing from the
    // MERGE_CELLS exemption a few lines above it in conversation-engine.service.ts).
    UNMERGE_CELLS: { type: 'UNMERGE_CELLS', sheetName: 'S', row: 0, col: 0, rowCount: 31, colCount: 9 },
    CLEAR_CONTENT: { type: 'CLEAR_CONTENT', sheetName: 'S', row: 1, col: 0, rowCount: 30, colCount: 9 },
    CLEAR_FORMAT: { type: 'CLEAR_FORMAT', sheetName: 'S', row: 1, col: 0, rowCount: 30, colCount: 9 },
    SORT_RANGE: { type: 'SORT_RANGE', sheetName: 'S', range: 'A1:I31', key: 1, ascending: true, hasHeaders: true },
    MOVE_RANGE: { type: 'MOVE_RANGE', sourceSheet: 'S', sourceRange: 'A1:B2', destSheet: 'S', destStartCell: 'D1' },
    COPY_FILTERED_RANGE: { type: 'COPY_FILTERED_RANGE', sourceSheet: 'S', sourceRange: 'A1:I31', hasHeaders: true, destSheet: 'T', destStartCell: 'A1', mode: 'copy' },
    FORMAT_MATCHING_ROWS: { type: 'FORMAT_MATCHING_ROWS', sheetName: 'S', range: 'A1:I31', hasHeaders: true, filter: { column: 'Amount', operator: 'greaterThan', value: 100 }, format: { fillColor: '#FF0000' } },
    SET_MATCHING_ROWS: { type: 'SET_MATCHING_ROWS', sheetName: 'S', range: 'A1:I31', hasHeaders: true, targetColumn: 'Status', value: 'Flagged' },
    DELETE_MATCHING_ROWS: { type: 'DELETE_MATCHING_ROWS', sheetName: 'S', range: 'A1:I31', hasHeaders: true },
    SET_ROW_HEIGHT: { type: 'SET_ROW_HEIGHT', sheetName: 'S', row: 1, rowCount: 4, height: 25 },
    CONDITIONAL_FORMAT: { type: 'CONDITIONAL_FORMAT', sheetName: 'S', range: 'E2:E31', rule: { kind: 'colorScale', colors: ['#63BE7B', '#F8696B'] } },
    AGGREGATE_TABLE: { type: 'AGGREGATE_TABLE', sourceSheet: 'S', sourceRange: 'A1:I31', groupByColumn: 'Supplier', aggregations: [{ column: 'Amount', fn: 'sum' }], destSheet: 'T', destStartCell: 'A1' },
    AUTO_FILTER: { type: 'AUTO_FILTER', sheetName: 'S', range: 'A1:I31' },
    FREEZE_PANES: { type: 'FREEZE_PANES', sheetName: 'S', freezeRows: 1 },
    UNFREEZE_PANES: { type: 'UNFREEZE_PANES', sheetName: 'S' },
    AUTOFIT_COLUMNS: { type: 'AUTOFIT_COLUMNS', sheetName: 'S' },
    HIDE_GRIDLINES: { type: 'HIDE_GRIDLINES', sheetName: 'S' },
    DATA_VALIDATION: { type: 'DATA_VALIDATION', sheetName: 'S', range: 'F2:F31', validation: { kind: 'list', listSource: ['IGST', 'Exempt'] } },
    DEFINE_NAMED_RANGE: { type: 'DEFINE_NAMED_RANGE', sheetName: 'S', name: 'TaxableAmount', range: 'E2:E31' },
    ADD_SHEET: { type: 'ADD_SHEET', name: 'New' },
    CREATE_SHEET: { type: 'CREATE_SHEET', sheetName: 'New' },
    DELETE_SHEET: { type: 'DELETE_SHEET', sheetName: 'Old' },
    RENAME_SHEET: { type: 'RENAME_SHEET', sheetName: 'S', newName: 'S2' },
    COPY_SHEET: { type: 'COPY_SHEET', sheetName: 'S', newSheetName: 'S copy' },
    MOVE_SHEET: { type: 'MOVE_SHEET', sheetName: 'S', position: 0 },
    HIDE_SHEET: { type: 'HIDE_SHEET', sheetName: 'S' },
    SHOW_SHEET: { type: 'SHOW_SHEET', sheetName: 'S' },
    SET_SHEET_COLOR: { type: 'SET_SHEET_COLOR', sheetName: 'S', color: '#0000FF' },
    CREATE_TABLE: { type: 'CREATE_TABLE', sheetName: 'S', range: 'A1:I31' },
    CREATE_CHART: { type: 'CREATE_CHART', sheetName: 'S', sourceSheetName: 'S', sourceRange: 'K1:L6', chartType: 'BarClustered', startCell: 'N2' },
    UPDATE_CHART: { type: 'UPDATE_CHART', sheetName: 'S', chartId: 'c1', title: 'New title' },
    ADD_COMMENT: { type: 'ADD_COMMENT', sheetName: 'S', row: 8, col: 4, comment: 'Check' },
  } as unknown as Record<string, SheetActionPayload>;

  it('has a sample for every advertised type (so this test cannot go vacuous)', () => {
    const missing = EXECUTOR_ADVERTISED_ACTION_TYPES.filter((type) => !SAMPLES[type]);
    expect(missing).toEqual([]);
  });

  it.each(EXECUTOR_ADVERTISED_ACTION_TYPES)('%s survives finalizeActions', (type) => {
    const sample = SAMPLES[type];
    const finalized = engine.finalizeActions([sample], analysis, undefined, 'do the thing');
    expect(finalized.map((action) => action.type)).toContain(type);
  });

  /**
   * TASKS.md #258 — the it.each above only checks the type survived, not that
   * the RIGHT fields came out — which is exactly how FILL_DOWN's bug hid:
   * it "survived" with row/col attached, fields the client handler silently
   * ignores, so the action looked fine here and did nothing real in Excel.
   */
  it('derives sourceRange/targetRange from a single range string (the shape models actually emit for "copy X down to Y")', () => {
    const finalized = engine.finalizeActions(
      [{ type: 'FILL_DOWN', sheetName: 'S', range: 'J2:J61' } as never],
      analysis,
      undefined,
      'Copy the formula in J2 down to J61',
    );
    expect(finalized).toEqual([
      expect.objectContaining({
        type: 'FILL_DOWN',
        sourceRange: 'J2',
        targetRange: 'J3:J61',
      }),
    ]);
  });

  it('leaves an already-correct sourceRange/targetRange pair untouched', () => {
    const finalized = engine.finalizeActions(
      [{ type: 'FILL_DOWN', sheetName: 'S', sourceRange: 'J2', targetRange: 'J3:J61' } as never],
      analysis,
      undefined,
      'do the thing',
    );
    expect(finalized).toEqual([
      expect.objectContaining({ sourceRange: 'J2', targetRange: 'J3:J61' }),
    ]);
  });
});
