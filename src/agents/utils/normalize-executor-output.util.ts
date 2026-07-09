import { SheetActionPayload, SheetActionType } from '../../excel-ai/types/sheet-actions.types';
import { Action, ExecutorOutput, SubTask } from '../types/agent.types';

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
  'DEFINE_NAMED_RANGE',
  'AUTOFIT_COLUMNS',
  'CLARIFY',
  'CHECKPOINT',
  'ADD_SHEET',
  'SORT_RANGE',
]);

function normalizeActionType(raw: unknown): SheetActionType | null {
  if (typeof raw !== 'string') return null;
  const upper = raw.trim().toUpperCase().replace(/[\s-]+/g, '_');
  if (upper === 'SORT' || upper === 'SORT_COLUMN' || upper === 'SORT_ROWS') {
    return 'SORT_RANGE';
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
  if (typeof record.oldName === 'string') action.oldName = record.oldName;
  if (typeof record.newName === 'string') action.newName = record.newName;
  if (typeof record.name === 'string') action.name = record.name;
  if (typeof record.tableName === 'string') action.tableName = record.tableName;
  if (typeof record.question === 'string') action.question = record.question;
  if (Array.isArray(record.options)) action.options = record.options.map(String);
  if (typeof record.message === 'string') action.message = record.message;
  if (typeof record.beforeColumn === 'string') action.beforeColumn = record.beforeColumn;
  if (Array.isArray(record.columns)) action.columns = record.columns.map(String);
  if (typeof record.afterRow === 'number') action.afterRow = record.afterRow;

  if (Array.isArray(record.data)) action.data = record.data;
  if (Array.isArray(record.values)) action.values = record.values;
  if (Array.isArray(record.data) && !Array.isArray(record.values)) {
    action.values = record.data;
  }
  if (Array.isArray(record.values) && type === 'ADD_ROW' && !Array.isArray(record.data)) {
    action.data = record.values;
  }

  if (Array.isArray(record.operations)) action.operations = record.operations as SheetActionPayload['operations'];
  if (Array.isArray(record.rows)) {
    if (record.rows.every((n) => typeof n === 'number')) {
      action.rowNumbers = record.rows as number[];
    }
  }
  if (Array.isArray(record.rowNumbers)) action.rowNumbers = record.rowNumbers as number[];

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

  if (record.format && typeof record.format === 'object') {
    action.format = record.format as SheetActionPayload['format'];
  }

  return action;
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
    subtaskId: String(parsed.subtaskId ?? subtask.id),
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
