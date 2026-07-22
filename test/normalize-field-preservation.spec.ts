import {
  normalizeExecutorOutput,
  normalizeSingleAction,
} from '../src/agents/utils/normalize-executor-output.util';
import { normalizeChartColorScheme } from '../src/agents/utils/chart-color-scheme.util';
import { SheetActionType } from '../src/excel-ai/types/sheet-actions.types';
import { SubTask } from '../src/agents/types/agent.types';

const subtask: SubTask = {
  id: 's1',
  description: 'test',
  targetSheet: 'Dashboard',
  dependsOn: [],
  estimatedActions: 1,
};

/** Every SheetActionType with representative optional fields populated. */
const FIELD_FIXTURES: Record<SheetActionType, Record<string, unknown>> = {
  SET_CELL: {
    type: 'SET_CELL',
    sheetName: 'Dashboard',
    row: 1,
    col: 2,
    value: 'x',
    explicitOverwriteConfirmed: true,
  },
  CLEAR_CELL: { type: 'CLEAR_CELL', sheetName: 'Dashboard', row: 0, col: 0 },
  HIGHLIGHT_CELL: {
    type: 'HIGHLIGHT_CELL',
    sheetName: 'Dashboard',
    row: 0,
    col: 0,
    color: '#FFFF00',
  },
  SET_FORMULA: {
    type: 'SET_FORMULA',
    sheetName: 'Dashboard',
    row: 1,
    col: 1,
    formula: '=A1+1',
  },
  ADD_ROW: { type: 'ADD_ROW', sheetName: 'Dashboard', data: ['a', 'b'], values: ['a', 'b'] },
  DELETE_ROW: { type: 'DELETE_ROW', sheetName: 'Dashboard', row: 2, rowNumbers: [2, 3] },
  INSERT_ROW: { type: 'INSERT_ROW', sheetName: 'Dashboard', row: 1, count: 1, position: 'above' },
  INSERT_COLUMN: {
    type: 'INSERT_COLUMN',
    sheetName: 'Dashboard',
    columnName: 'Net',
    position: 'afterLastColumn',
    afterColumn: 'Amount',
    formula: '=A{row}',
  },
  DELETE_COLUMN: { type: 'DELETE_COLUMN', sheetName: 'Dashboard', col: 1 },
  HIDE_ROW: { type: 'HIDE_ROW', sheetName: 'Dashboard', row: 1 },
  UNHIDE_ROW: { type: 'UNHIDE_ROW', sheetName: 'Dashboard', row: 1 },
  SHOW_ROW: { type: 'SHOW_ROW', sheetName: 'Dashboard', row: 1 },
  HIDE_COLUMN: { type: 'HIDE_COLUMN', sheetName: 'Dashboard', col: 1 },
  UNHIDE_COLUMN: { type: 'UNHIDE_COLUMN', sheetName: 'Dashboard', col: 1 },
  SHOW_COLUMN: { type: 'SHOW_COLUMN', sheetName: 'Dashboard', col: 1 },
  SET_ROW_HEIGHT: { type: 'SET_ROW_HEIGHT', sheetName: 'Dashboard', row: 0, height: 20 },
  SET_COLUMN_WIDTH: { type: 'SET_COLUMN_WIDTH', sheetName: 'Dashboard', col: 0, width: 15 },
  FREEZE_PANES: { type: 'FREEZE_PANES', sheetName: 'Dashboard', freezeRows: 1, freezeColumns: 0 },
  UNFREEZE_PANES: { type: 'UNFREEZE_PANES', sheetName: 'Dashboard' },
  SET_ZOOM: { type: 'SET_ZOOM', sheetName: 'Dashboard', zoomPercent: 120 },
  PROTECT_SHEET: { type: 'PROTECT_SHEET', sheetName: 'Dashboard' },
  UNPROTECT_SHEET: { type: 'UNPROTECT_SHEET', sheetName: 'Dashboard' },
  MERGE_CELLS: {
    type: 'MERGE_CELLS',
    sheetName: 'Dashboard',
    range: 'A1:B1',
    row: 0,
    col: 0,
    rowCount: 1,
    colCount: 2,
    mergeAcross: false,
  },
  UNMERGE_CELLS: {
    type: 'UNMERGE_CELLS',
    sheetName: 'Dashboard',
    range: 'A1:B1',
    row: 0,
    col: 0,
    rowCount: 1,
    colCount: 2,
  },
  CLEAR_CONTENT: {
    type: 'CLEAR_CONTENT',
    sheetName: 'Dashboard',
    range: 'A1:B2',
    row: 0,
    col: 0,
    rowCount: 2,
    colCount: 2,
  },
  CLEAR_FORMAT: {
    type: 'CLEAR_FORMAT',
    sheetName: 'Dashboard',
    range: 'A1:B2',
    row: 0,
    col: 0,
    rowCount: 2,
    colCount: 2,
  },
  CLEAR_ALL: {
    type: 'CLEAR_ALL',
    sheetName: 'Dashboard',
    range: 'A1:B2',
    row: 0,
    col: 0,
    rowCount: 2,
    colCount: 2,
  },
  FORMAT_RANGE: {
    type: 'FORMAT_RANGE',
    sheetName: 'Dashboard',
    range: 'A1:L1',
    row: 0,
    col: 0,
    rowCount: 1,
    colCount: 12,
    format: { bold: true, fillColor: '#DDEEFF' },
  },
  FILL_DOWN: { type: 'FILL_DOWN', sheetName: 'Dashboard', row: 1, col: 0, endRow: 10 },
  FILL_RIGHT: { type: 'FILL_RIGHT', sheetName: 'Dashboard', row: 0, col: 0, endCol: 5 },
  CREATE_SHEET: { type: 'CREATE_SHEET', sheetName: 'New', name: 'New', relativeTo: 'Dashboard' },
  DELETE_SHEET: { type: 'DELETE_SHEET', sheetName: 'Old' },
  RENAME_SHEET: { type: 'RENAME_SHEET', oldName: 'Old', newName: 'New', sheetName: 'Old' },
  COPY_SHEET: {
    type: 'COPY_SHEET',
    sheetName: 'Dashboard',
    sourceName: 'Dashboard',
    newName: 'Copy',
    copyFrom: 'Dashboard',
  },
  HIDE_SHEET: { type: 'HIDE_SHEET', sheetName: 'Dashboard' },
  SHOW_SHEET: { type: 'SHOW_SHEET', sheetName: 'Dashboard' },
  SET_SHEET_COLOR: { type: 'SET_SHEET_COLOR', sheetName: 'Dashboard', color: '#00FF00' },
  ADD_COMMENT: {
    type: 'ADD_COMMENT',
    sheetName: 'Dashboard',
    row: 0,
    col: 0,
    comment: 'note',
  },
  DELETE_COMMENT: { type: 'DELETE_COMMENT', sheetName: 'Dashboard', row: 0, col: 0 },
  WRITE_TABLE: {
    type: 'WRITE_TABLE',
    sheetName: 'Dashboard',
    headers: ['A', 'B'],
    rows: [[1, 2]],
  },
  BATCH_SET: {
    type: 'BATCH_SET',
    sheetName: 'Dashboard',
    operations: [{ address: 'A1', value: 1 }],
  },
  CREATE_TABLE: {
    type: 'CREATE_TABLE',
    sheetName: 'Dashboard',
    tableName: 'T1',
    name: 'T1',
    range: 'A1:B10',
    hasHeaders: true,
    style: 'TableStyleMedium2',
  },
  CREATE_CHART: {
    type: 'CREATE_CHART',
    sheetName: 'Dashboard',
    sourceSheetName: 'Dashboard',
    sourceRange: 'A1:B10',
    chartType: 'BarClustered',
    title: 'Spend',
    startCell: 'D2',
    endCell: 'K16',
    destCell: 'D2',
    colorScheme: 'green',
    chartId: 'Chart_spend',
  },
  DEFINE_NAMED_RANGE: {
    type: 'DEFINE_NAMED_RANGE',
    sheetName: 'Dashboard',
    name: 'MyRange',
    formula: "='Dashboard'!$A$1:$B$10",
    comment: 'kpi',
  },
  AUTOFIT_COLUMNS: { type: 'AUTOFIT_COLUMNS', sheetName: 'Dashboard', columns: ['A', 'B'] },
  CLARIFY: {
    type: 'CLARIFY',
    question: 'Which column?',
    options: ['A', 'B'],
    message: 'need input',
  },
  CHECKPOINT: { type: 'CHECKPOINT', message: 'step done' },
  ADD_SHEET: { type: 'ADD_SHEET', name: 'SheetX', sheetName: 'SheetX', position: 'after' },
  SORT_RANGE: {
    type: 'SORT_RANGE',
    sheetName: 'Dashboard',
    range: 'A1:B10',
    key: 1,
    ascending: false,
    hasHeaders: true,
    columnName: 'Amount',
  },
  COPY_FILTERED_RANGE: {
    type: 'COPY_FILTERED_RANGE',
    sourceSheet: 'Purchase Register',
    sourceRange: 'A1:L200',
    destSheet: 'Pending',
    destStartCell: 'A1',
    hasHeaders: true,
    mode: 'copy',
    filter: { column: 'Status', operator: 'equals', value: 'Pending' },
    explicitOverwriteConfirmed: true,
  },
  FORMAT_MATCHING_ROWS: {
    type: 'FORMAT_MATCHING_ROWS',
    sheetName: 'Purchase Register',
    range: 'A1:L200',
    hasHeaders: true,
    filter: { column: 'Status', operator: 'equals', value: 'Open' },
    format: { fillColor: '#FFEEEE' },
  },
  MOVE_RANGE: {
    type: 'MOVE_RANGE',
    sourceSheet: 'Sheet1',
    sourceRange: 'A1:B5',
    destSheet: 'Sheet2',
    destStartCell: 'A1',
    mode: 'move',
  },
  AGGREGATE_TABLE: {
    type: 'AGGREGATE_TABLE',
    sourceSheet: 'Purchase Register',
    sourceRange: 'A1:L200',
    groupByColumn: 'Date',
    groupByTransform: 'month',
    aggregations: [{ column: 'Amount', fn: 'sum', outputLabel: 'Total' }],
    sortBy: { column: 'Total', direction: 'desc' },
    topN: 5,
    destSheet: 'Dashboard',
    destStartCell: 'A4',
    hasHeaders: true,
    explicitOverwriteConfirmed: true,
  },
  UPDATE_CHART: {
    type: 'UPDATE_CHART',
    sheetName: 'Dashboard',
    chartId: 'Chart_TotalTaxByDate',
    chartType: 'BarClustered',
    colorScheme: 'green',
  },
};

/** Fields that may be intentionally transformed or aliased during normalize. */
const INTENTIONAL_TRANSFORMS: Partial<Record<SheetActionType, string[]>> = {
  // range string may expand into indices; keep range too when already present
  FORMAT_RANGE: [],
  MERGE_CELLS: [],
  UNMERGE_CELLS: [],
  CLEAR_CONTENT: [],
  CLEAR_FORMAT: [],
  CLEAR_ALL: [],
  ADD_COMMENT: [],
  DELETE_COMMENT: [],
  // CREATE_CHART may alias range → sourceRange; chartId auto-filled if missing
  CREATE_CHART: [],
  // COPY may alias sheetName → sourceSheet
  COPY_FILTERED_RANGE: [],
  MOVE_RANGE: [],
  CREATE_TABLE: ['name'], // name mirrored into tableName
  ADD_ROW: [], // values/data mirrored
};

function deepEqualish(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

describe('normalize chart colorScheme (Spec 18 Bug 3)', () => {
  it('preserves green on UPDATE_CHART (exact repro)', () => {
    const raw = {
      type: 'UPDATE_CHART',
      sheetName: 'Dashboard',
      chartId: 'Chart_TotalTaxByDate',
      chartType: 'BarClustered',
      colorScheme: 'green',
    };
    const normalized = normalizeSingleAction(raw, 'Dashboard');
    expect(normalized).toMatchObject({
      type: 'UPDATE_CHART',
      sheetName: 'Dashboard',
      chartId: 'Chart_TotalTaxByDate',
      chartType: 'BarClustered',
      colorScheme: 'green',
    });
  });

  it('maps gray→grey and greeen→green aliases', () => {
    expect(normalizeChartColorScheme('gray')).toBe('grey');
    expect(normalizeChartColorScheme('greeen')).toBe('green');
    expect(normalizeChartColorScheme('Blue Grey')).toBe('blueGrey');
  });

  it('preserves green on CREATE_CHART', () => {
    const normalized = normalizeSingleAction(FIELD_FIXTURES.CREATE_CHART, 'Dashboard');
    expect(normalized?.colorScheme).toBe('green');
    expect(normalized?.chartId).toBe('Chart_spend');
  });
});

describe('normalize field-preservation regression (all SheetActionType)', () => {
  const types = Object.keys(FIELD_FIXTURES) as SheetActionType[];

  it('covers every SheetActionType from the type union', () => {
    // Sanity: fixture map must include all known types used by normalize KNOWN_TYPES
    expect(types.length).toBeGreaterThan(40);
    expect(types).toContain('UPDATE_CHART');
    expect(types).toContain('AGGREGATE_TABLE');
    expect(types).toContain('COPY_FILTERED_RANGE');
    expect(types).toContain('FORMAT_RANGE');
  });

  it.each(types)('preserves optional fields for %s', (type) => {
    const raw = FIELD_FIXTURES[type];
    const normalized = normalizeSingleAction(raw, 'Dashboard');
    expect(normalized).not.toBeNull();
    expect(normalized!.type).toBe(type);

    const intentional = new Set(INTENTIONAL_TRANSFORMS[type] ?? []);
    for (const [key, value] of Object.entries(raw)) {
      if (key === 'type') continue;
      if (intentional.has(key)) continue;
      const actual = (normalized as unknown as Record<string, unknown>)[key];
      expect(actual).toBeDefined();
      if (typeof value === 'object' && value !== null) {
        expect(deepEqualish(actual, value)).toBe(true);
      } else {
        expect(actual).toEqual(value);
      }
    }
  });

  it('normalizeExecutorOutput keeps UPDATE_CHART colorScheme through the public API', () => {
    const result = normalizeExecutorOutput(
      {
        subtaskId: 's1',
        actions: [FIELD_FIXTURES.UPDATE_CHART],
        isDone: true,
      },
      subtask,
    );
    expect(result.actions[0]).toMatchObject({
      type: 'UPDATE_CHART',
      colorScheme: 'green',
      chartId: 'Chart_TotalTaxByDate',
    });
  });
});
