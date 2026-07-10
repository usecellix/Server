export type CellValue = string | number | boolean | null;

export interface WorkbookContext {
  activeSheet: string;
  sheets: string[];
  headers: Record<string, string>;
  dataRange: string;
  lastDataRow: number;
  headerRow: number;
  dataStartRow: number;
  totalRows: number;
  columnCount: number;
  columnLetters: string[];
}

export interface FormatSpec {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  horizontalAlignment?: 'left' | 'center' | 'right';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numberFormat?: string;
  borders?: 'all' | 'outer' | 'bottom' | 'none';
}

export interface RangeRef {
  row: number;
  col: number;
  rowCount?: number;
  colCount?: number;
}

export type SheetActionType =
  | 'SET_CELL'
  | 'CLEAR_CELL'
  | 'HIGHLIGHT_CELL'
  | 'SET_FORMULA'
  | 'ADD_ROW'
  | 'DELETE_ROW'
  | 'INSERT_ROW'
  | 'INSERT_COLUMN'
  | 'DELETE_COLUMN'
  | 'HIDE_ROW'
  | 'UNHIDE_ROW'
  | 'SHOW_ROW'
  | 'HIDE_COLUMN'
  | 'UNHIDE_COLUMN'
  | 'SHOW_COLUMN'
  | 'SET_ROW_HEIGHT'
  | 'SET_COLUMN_WIDTH'
  | 'FREEZE_PANES'
  | 'UNFREEZE_PANES'
  | 'SET_ZOOM'
  | 'PROTECT_SHEET'
  | 'UNPROTECT_SHEET'
  | 'MERGE_CELLS'
  | 'UNMERGE_CELLS'
  | 'CLEAR_CONTENT'
  | 'CLEAR_FORMAT'
  | 'CLEAR_ALL'
  | 'FORMAT_RANGE'
  | 'FILL_DOWN'
  | 'FILL_RIGHT'
  | 'CREATE_SHEET'
  | 'DELETE_SHEET'
  | 'RENAME_SHEET'
  | 'COPY_SHEET'
  | 'HIDE_SHEET'
  | 'SHOW_SHEET'
  | 'SET_SHEET_COLOR'
  | 'ADD_COMMENT'
  | 'DELETE_COMMENT'
  | 'WRITE_TABLE'
  | 'BATCH_SET'
  | 'CREATE_TABLE'
  | 'DEFINE_NAMED_RANGE'
  | 'AUTOFIT_COLUMNS'
  | 'CLARIFY'
  | 'CHECKPOINT'
  | 'ADD_SHEET'
  | 'SORT_RANGE';

export interface BatchSetOperation {
  address: string;
  value?: unknown;
  formula?: string;
  format?: FormatSpec;
}

export interface SheetActionPayload {
  type: SheetActionType;
  row?: number;
  col?: number;
  rowCount?: number;
  colCount?: number;
  value?: unknown;
  color?: string;
  formula?: string;
  data?: unknown[];
  count?: number;
  position?: 'above' | 'below' | 'left' | 'right' | 'before' | 'after';
  height?: number;
  width?: number;
  freezeRows?: number;
  freezeColumns?: number;
  zoomPercent?: number;
  mergeAcross?: boolean;
  format?: FormatSpec;
  endRow?: number;
  endCol?: number;
  sheetName?: string;
  newSheetName?: string;
  relativeTo?: string;
  comment?: string;
  headers?: string[];
  rows?: unknown[][];
  afterRow?: number;
  values?: unknown[];
  address?: string;
  range?: string;
  sourceRange?: string;
  targetRange?: string;
  operations?: BatchSetOperation[];
  rowNumbers?: number[];
  beforeColumn?: string;
  columns?: string[];
  oldName?: string;
  sourceName?: string;
  newName?: string;
  name?: string;
  tableName?: string;
  hasHeaders?: boolean;
  style?: string;
  question?: string;
  options?: string[];
  message?: string;
  copyFrom?: string;
  /** SORT_RANGE — 0-based column index within range */
  key?: number;
  ascending?: boolean;
  columnName?: string;
}

export type SheetAction = SheetActionPayload;

export type IntentType = 'ACTION' | 'EXPLAIN' | 'FIX' | 'DATA_QUESTION' | 'FORMULA_HELP';

/** User-selected operational mode from the add-in. */
export type AssistantMode = 'ask' | 'action' | 'plan';

export interface IntentClassification {
  intent: IntentType;
  confidence: 'high' | 'medium' | 'low';
  subIntent?: string;
}
