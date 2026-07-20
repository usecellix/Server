import { SheetActionPayload, SheetActionType } from '../../excel-ai/types/sheet-actions.types';
import { Action, ExecutorOutput, SubTask } from '../types/agent.types';
import { stripSheetPrefix } from './range-address.util';
import { parseA1Range } from './range-merge.util';

/** Action types that sanitizeAction requires integer row/col for. */
const INDEX_RANGE_ACTION_TYPES = new Set<SheetActionType>([
  'FORMAT_RANGE',
  'MERGE_CELLS',
  'UNMERGE_CELLS',
  'CLEAR_CONTENT',
  'CLEAR_FORMAT',
  'CLEAR_ALL',
  'ADD_COMMENT',
  'DELETE_COMMENT',
]);

const KNOWN_TYPES = new Set<string>([
  'SET_CELL',
  'CLEAR_CELL',
  'HIGHLIGHT_CELL',
  'SET_FORMULA',
  'ADD_ROW',
  'DELETE_ROW',
  'INSERT_ROW',
  'INSERT_COLUMN',
  'DELETE_COLUMN',
  'HIDE_ROW',
  'UNHIDE_ROW',
  'SHOW_ROW',
  'HIDE_COLUMN',
  'UNHIDE_COLUMN',
  'SHOW_COLUMN',
  'SET_ROW_HEIGHT',
  'SET_COLUMN_WIDTH',
  'FREEZE_PANES',
  'UNFREEZE_PANES',
  'SET_ZOOM',
  'PROTECT_SHEET',
  'UNPROTECT_SHEET',
  'MERGE_CELLS',
  'UNMERGE_CELLS',
  'CLEAR_CONTENT',
  'CLEAR_FORMAT',
  'CLEAR_ALL',
  'FORMAT_RANGE',
  'FILL_DOWN',
  'FILL_RIGHT',
  'CREATE_SHEET',
  'DELETE_SHEET',
  'RENAME_SHEET',
  'COPY_SHEET',
  'HIDE_SHEET',
  'SHOW_SHEET',
  'SET_SHEET_COLOR',
  'ADD_COMMENT',
  'DELETE_COMMENT',
  'WRITE_TABLE',
  'BATCH_SET',
  'CREATE_TABLE',
  'CREATE_CHART',
  'DEFINE_NAMED_RANGE',
  'AUTOFIT_COLUMNS',
  'CLARIFY',
  'CHECKPOINT',
  'ADD_SHEET',
  'SORT_RANGE',
  'COPY_FILTERED_RANGE',
  'FORMAT_MATCHING_ROWS',
  'MOVE_RANGE',
  'AGGREGATE_TABLE',
  'UPDATE_CHART',
]);

function normalizeActionType(raw: unknown): SheetActionType | null {
  if (typeof raw !== 'string') return null;
  const upper = raw.trim().toUpperCase().replace(/[\s-]+/g, '_');
  if (upper === 'SORT' || upper === 'SORT_COLUMN' || upper === 'SORT_ROWS') {
    return 'SORT_RANGE';
  }
  if (
    upper === 'COPY_FILTERED' ||
    upper === 'COPY_FILTER' ||
    upper === 'FILTER_COPY' ||
    upper === 'COPY_ROWS'
  ) {
    return 'COPY_FILTERED_RANGE';
  }
  if (upper === 'MOVE' || upper === 'MOVE_ROWS') {
    return 'MOVE_RANGE';
  }
  if (
    upper === 'FORMAT_MATCHING' ||
    upper === 'FORMAT_MATCHING_ROWS' ||
    upper === 'HIGHLIGHT_MATCHING' ||
    upper === 'HIGHLIGHT_ROWS' ||
    upper === 'CONDITIONAL_FORMAT'
  ) {
    return 'FORMAT_MATCHING_ROWS';
  }
  if (KNOWN_TYPES.has(upper)) {
    return upper as SheetActionType;
  }
  return null;
}

function normalizeSingleAction(
  raw: unknown,
  defaultSheetName: string,
): SheetActionPayload | null {
  if (!raw || typeof raw !== 'object') return null;
  const record = raw as Record<string, unknown>;
  const type = normalizeActionType(record.type);
  if (!type) return null;

  const action: SheetActionPayload = { type };

  if (typeof record.sheetName === 'string') action.sheetName = record.sheetName;
  else if (defaultSheetName) action.sheetName = defaultSheetName;

  if (record.row !== undefined) action.row = Number(record.row);
  if (record.col !== undefined) action.col = Number(record.col);
  if (record.rowCount !== undefined) action.rowCount = Number(record.rowCount);
  if (record.colCount !== undefined) action.colCount = Number(record.colCount);
  if (record.zoomPercent !== undefined) action.zoomPercent = Number(record.zoomPercent);
  if (record.value !== undefined) action.value = record.value;
  if (typeof record.formula === 'string') action.formula = record.formula;
  if (typeof record.address === 'string') action.address = record.address;
  if (typeof record.range === 'string') action.range = record.range;
  if (typeof record.sourceRange === 'string') action.sourceRange = record.sourceRange;
  if (typeof record.targetRange === 'string') action.targetRange = record.targetRange;
  if (typeof record.sourceSheetName === 'string') {
    action.sourceSheetName = record.sourceSheetName;
  }
  if (typeof record.chartType === 'string') action.chartType = record.chartType;
  if (typeof record.title === 'string') action.title = record.title;
  if (typeof record.startCell === 'string') action.startCell = record.startCell;
  if (typeof record.endCell === 'string') action.endCell = record.endCell;
  if (typeof record.oldName === 'string') action.oldName = record.oldName;
  if (typeof record.newName === 'string') action.newName = record.newName;
  if (typeof record.name === 'string') action.name = record.name;
  if (typeof record.tableName === 'string') action.tableName = record.tableName;
  if (typeof record.question === 'string') action.question = record.question;
  if (Array.isArray(record.options)) action.options = record.options.map(String);
  if (typeof record.message === 'string') action.message = record.message;
  if (typeof record.beforeColumn === 'string') action.beforeColumn = record.beforeColumn;
  if (typeof record.afterColumn === 'string') action.afterColumn = record.afterColumn;
  if (Array.isArray(record.columns)) action.columns = record.columns.map(String);
  if (typeof record.afterRow === 'number') action.afterRow = record.afterRow;
  if (record.explicitOverwriteConfirmed === true) {
    action.explicitOverwriteConfirmed = true;
  }

  if (type === 'INSERT_COLUMN') {
    if (typeof record.columnName === 'string') action.columnName = record.columnName;
    if (record.position === 'afterLastColumn') {
      action.position = 'afterLastColumn';
    } else if (record.position && typeof record.position === 'object') {
      const pos = record.position as Record<string, unknown>;
      if (typeof pos.afterColumn === 'string') {
        action.afterColumn = pos.afterColumn;
      }
    } else if (
      record.position === 'above' ||
      record.position === 'below' ||
      record.position === 'left' ||
      record.position === 'right' ||
      record.position === 'before' ||
      record.position === 'after'
    ) {
      action.position = record.position;
    }
    if (typeof record.formula === 'string') action.formula = record.formula;
  }

  if (Array.isArray(record.data)) action.data = record.data;
  if (Array.isArray(record.values)) action.values = record.values;
  if (Array.isArray(record.data) && !Array.isArray(record.values)) {
    action.values = record.data;
  }
  if (Array.isArray(record.values) && type === 'ADD_ROW' && !Array.isArray(record.data)) {
    action.data = record.values;
  }

  if (Array.isArray(record.operations)) action.operations = record.operations as SheetActionPayload['operations'];
  if (Array.isArray(record.headers)) {
    action.headers = record.headers.map((h) => (h == null ? '' : String(h)));
  }
  if (Array.isArray(record.rows)) {
    if (record.rows.every((n) => typeof n === 'number')) {
      action.rowNumbers = record.rows as number[];
    } else {
      // WRITE_TABLE / create_data payloads use 2D row arrays
      action.rows = record.rows as unknown[][];
    }
  }
  if (Array.isArray(record.rowNumbers)) action.rowNumbers = record.rowNumbers as number[];

  if (type === 'CREATE_TABLE') {
    const tableName = record.tableName ?? record.name;
    if (typeof tableName === 'string') action.tableName = tableName.trim();
    action.hasHeaders =
      record.hasHeaders === undefined ? true : Boolean(record.hasHeaders);
  }

  if (type === 'SORT_RANGE') {
    const key =
      record.key ??
      record.columnIndex ??
      record.colIndex ??
      record.col ??
      record.column;
    if (key !== undefined) action.key = Number(key);
    if (record.ascending !== undefined) action.ascending = Boolean(record.ascending);
    if (record.hasHeaders !== undefined) action.hasHeaders = Boolean(record.hasHeaders);
    if (typeof record.columnName === 'string') action.columnName = record.columnName;
  }

  if (type === 'COPY_FILTERED_RANGE' || type === 'MOVE_RANGE') {
    if (typeof record.sourceSheet === 'string') action.sourceSheet = record.sourceSheet;
    else if (typeof record.sheetName === 'string') action.sourceSheet = record.sheetName;
    if (typeof record.sourceRange === 'string') action.sourceRange = record.sourceRange;
    else if (typeof record.range === 'string') action.sourceRange = record.range;
    if (typeof record.destSheet === 'string') action.destSheet = record.destSheet;
    else if (typeof record.targetSheet === 'string') action.destSheet = record.targetSheet;
    if (typeof record.destStartCell === 'string') action.destStartCell = record.destStartCell;
    else if (typeof record.startCell === 'string') action.destStartCell = record.startCell;
    if (record.hasHeaders !== undefined) action.hasHeaders = Boolean(record.hasHeaders);
    else if (type === 'COPY_FILTERED_RANGE') action.hasHeaders = true;
    if (record.mode === 'copy' || record.mode === 'move') {
      action.mode = record.mode;
    } else if (type === 'COPY_FILTERED_RANGE') {
      action.mode = 'copy';
    }
    if (record.filter && typeof record.filter === 'object') {
      const filter = record.filter as Record<string, unknown>;
      if (
        typeof filter.column === 'string' &&
        typeof filter.operator === 'string' &&
        (typeof filter.value === 'string' || typeof filter.value === 'number')
      ) {
        action.filter = {
          column: filter.column,
          operator: filter.operator as NonNullable<SheetActionPayload['filter']>['operator'],
          value: filter.value,
        };
      }
    }
  }

  if (type === 'FORMAT_MATCHING_ROWS') {
    if (typeof record.sheetName === 'string') action.sheetName = record.sheetName;
    if (typeof record.range === 'string') action.range = stripSheetPrefix(record.range);
    else if (typeof record.sourceRange === 'string') {
      action.range = stripSheetPrefix(record.sourceRange);
    }
    if (record.hasHeaders !== undefined) action.hasHeaders = Boolean(record.hasHeaders);
    else action.hasHeaders = true;
    if (record.filter && typeof record.filter === 'object') {
      const filter = record.filter as Record<string, unknown>;
      if (
        typeof filter.column === 'string' &&
        typeof filter.operator === 'string' &&
        (typeof filter.value === 'string' || typeof filter.value === 'number')
      ) {
        action.filter = {
          column: filter.column,
          operator: filter.operator as NonNullable<SheetActionPayload['filter']>['operator'],
          value: filter.value,
        };
      }
    }
    if (record.format && typeof record.format === 'object') {
      action.format = record.format as SheetActionPayload['format'];
    } else if (typeof record.color === 'string') {
      action.format = { fillColor: record.color };
    }
    if (
      action.format &&
      typeof action.format === 'object' &&
      (action.format as { clearFill?: boolean }).clearFill !== true &&
      !action.format.fillColor
    ) {
      const formatRecord = record.format as Record<string, unknown> | undefined;
      if (formatRecord?.clearFill === true) {
        action.format = { clearFill: true };
      }
    }
  }

  if (type === 'AGGREGATE_TABLE') {
    if (typeof record.sourceSheet === 'string') action.sourceSheet = record.sourceSheet;
    else if (typeof record.sheetName === 'string') action.sourceSheet = record.sheetName;
    if (typeof record.sourceRange === 'string') action.sourceRange = record.sourceRange;
    else if (typeof record.range === 'string') action.sourceRange = record.range;
    if (typeof record.groupByColumn === 'string') action.groupByColumn = record.groupByColumn;
    if (typeof record.destSheet === 'string') action.destSheet = record.destSheet;
    if (typeof record.destStartCell === 'string') action.destStartCell = record.destStartCell;
    else if (typeof record.startCell === 'string') action.destStartCell = record.startCell;
    if (record.hasHeaders !== undefined) action.hasHeaders = Boolean(record.hasHeaders);
    else action.hasHeaders = true;
    if (typeof record.topN === 'number') action.topN = record.topN;
    if (record.sortBy && typeof record.sortBy === 'object') {
      const sortBy = record.sortBy as Record<string, unknown>;
      if (typeof sortBy.column === 'string') {
        action.sortBy = {
          column: sortBy.column,
          direction: sortBy.direction === 'desc' ? 'desc' : 'asc',
        };
      }
    }
    if (Array.isArray(record.aggregations)) {
      action.aggregations = record.aggregations
        .filter((item): item is Record<string, unknown> => Boolean(item) && typeof item === 'object')
        .map((item) => ({
          column: String(item.column ?? ''),
          fn: (['sum', 'count', 'average', 'max', 'min'].includes(String(item.fn))
            ? String(item.fn)
            : 'sum') as NonNullable<SheetActionPayload['aggregations']>[number]['fn'],
          outputLabel: String(item.outputLabel ?? item.column ?? 'Value'),
        }))
        .filter((item) => item.column);
    }
  }

  if (type === 'CREATE_CHART') {
    if (typeof record.sourceSheetName === 'string') action.sourceSheetName = record.sourceSheetName;
    else if (typeof record.sheetName === 'string') action.sourceSheetName = record.sheetName;
    if (typeof record.sourceRange === 'string') action.sourceRange = record.sourceRange;
    else if (typeof record.range === 'string') action.sourceRange = record.range;
    if (typeof record.chartType === 'string') action.chartType = record.chartType;
    if (typeof record.title === 'string') action.title = record.title;
    if (typeof record.startCell === 'string') action.startCell = record.startCell;
    if (typeof record.endCell === 'string') action.endCell = record.endCell;
    if (typeof record.destCell === 'string') action.destCell = record.destCell;
    if (typeof record.chartId === 'string') action.chartId = record.chartId;
    else action.chartId = `Chart_${Date.now().toString(36)}`;
    if (
      record.colorScheme === 'default' ||
      record.colorScheme === 'blue' ||
      record.colorScheme === 'grey' ||
      record.colorScheme === 'blueGrey'
    ) {
      action.colorScheme = record.colorScheme;
    }
  }

  if (type === 'UPDATE_CHART') {
    if (typeof record.chartId === 'string') action.chartId = record.chartId;
    if (typeof record.chartType === 'string') action.chartType = record.chartType;
    if (
      record.colorScheme === 'default' ||
      record.colorScheme === 'blue' ||
      record.colorScheme === 'grey' ||
      record.colorScheme === 'blueGrey'
    ) {
      action.colorScheme = record.colorScheme;
    }
  }

  if (record.format && typeof record.format === 'object') {
    action.format = record.format as SheetActionPayload['format'];
  }

  // sanitizeAction requires integer row/col for these types. Models often emit an
  // A1 `range` string instead — convert when possible so sanitize does not drop them.
  if (INDEX_RANGE_ACTION_TYPES.has(type)) {
    expandRangeStringToIndices(action);
  }

  return action;
}

function isValidIndex(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0;
}

/**
 * If row/col are already valid, leave as-is. Otherwise parse `range` (e.g. "A1:L1")
 * into 0-indexed row/col/rowCount/colCount. Malformed actions (no range, no indices)
 * are left unchanged for sanitizeAction to reject.
 */
function expandRangeStringToIndices(action: SheetActionPayload): void {
  if (isValidIndex(action.row) && isValidIndex(action.col)) {
    return;
  }
  if (typeof action.range !== 'string' || !action.range.trim()) {
    return;
  }
  // Sheet prefix is optional noise; A1 alone is enough to resolve indices.
  const parsed = parseA1Range(stripSheetPrefix(action.range));
  if (!parsed) {
    return;
  }
  action.row = parsed.startRow;
  action.col = parsed.startCol;
  action.rowCount = parsed.endRow - parsed.startRow + 1;
  action.colCount = parsed.endCol - parsed.startCol + 1;
}

export function normalizeExecutorOutput(
  parsed: Record<string, unknown>,
  subtask: SubTask,
): ExecutorOutput {
  const rawActions = Array.isArray(parsed.actions) ? parsed.actions : [];
  const actions: Action[] = rawActions
    .map((a) => normalizeSingleAction(a, subtask.targetSheet))
    .filter((a): a is Action => a !== null);

  const isDone =
    parsed.isDone === true ||
    parsed.isDone === false
      ? Boolean(parsed.isDone)
      : actions.length > 0;

  const nextStep =
    typeof parsed.nextStep === 'string'
      ? parsed.nextStep
      : typeof parsed.message === 'string'
        ? parsed.message
        : undefined;

  const toolRequest = parseToolRequest(parsed.toolRequest);

  return {
    // The executor is invoked for one known subtask. Canonicalize the ID instead
    // of letting an omitted, empty, or invented model value break loop progress.
    subtaskId: subtask.id,
    actions,
    isDone,
    nextStep,
    toolRequest,
  };
}

function parseToolRequest(raw: unknown): ExecutorOutput['toolRequest'] {
  if (!raw || typeof raw !== 'object') return undefined;
  const record = raw as Record<string, unknown>;
  const name = String(record.name ?? record.tool ?? '').trim();
  if (name !== 'get_range_data') return undefined;
  const sheet = String(record.sheet ?? record.sheetName ?? '').trim();
  const range = String(record.range ?? '').trim();
  if (!sheet || !range) return undefined;
  return { name: 'get_range_data', sheet, range };
}
