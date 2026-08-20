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
  /** When true, remove background fill from the range (leaves font/borders intact). */
  clearFill?: boolean;
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
  | 'AUTO_FILTER'
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
  | 'DELETE_TABLE'
  | 'CREATE_CHART'
  | 'DEFINE_NAMED_RANGE'
  | 'AUTOFIT_COLUMNS'
  | 'CLARIFY'
  | 'CHECKPOINT'
  | 'ADD_SHEET'
  | 'SORT_RANGE'
  | 'COPY_FILTERED_RANGE'
  | 'FORMAT_MATCHING_ROWS'
  | 'SET_MATCHING_ROWS'
  | 'MOVE_RANGE'
  | 'AGGREGATE_TABLE'
  | 'UPDATE_CHART'
  | 'DELETE_CHART'
  | 'CONDITIONAL_FORMAT'
  | 'DELETE_CONDITIONAL_FORMAT';

/** Excel's own cell-value conditional-format comparison operators (Office.js `ConditionalCellValueOperator`). */
export type ConditionalFormatOperator =
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equalTo'
  | 'notEqualTo'
  | 'between'
  | 'notBetween';

/** Single-column numeric-comparison variant (TASKS.md #32/#33). */
export interface ConditionalFormatCellValueRule {
  kind: 'cellValue';
  operator: ConditionalFormatOperator;
  value: number | string;
  /** Required when operator is 'between' / 'notBetween'. */
  value2?: number | string;
  format: FormatSpec;
}

/**
 * A boolean Excel formula, evaluated relative to the top-left cell of `range`
 * — e.g. `"=$C2<$B2*0.9"` for a cross-column comparison the cellValue variant
 * can't express. This is what VISION.md's own example needs: "highlight the
 * regions where revenue dropped more than 10%" (TASKS.md #35).
 */
export interface ConditionalFormatFormulaRule {
  kind: 'formula';
  formula: string;
  format: FormatSpec;
}

/**
 * Highlights the top or bottom N items (or N%) of `range` by value — e.g.
 * "highlight the top 5 suppliers by total". Re-ranks live as values change,
 * including a new row entering the top/bottom set (TASKS.md #36).
 */
export interface ConditionalFormatTopBottomRule {
  kind: 'topBottom';
  side: 'top' | 'bottom';
  /** Item count by default; a 0-100 percentage when `isPercent` is true. */
  rank: number;
  isPercent?: boolean;
  format: FormatSpec;
}

/**
 * A 2- or 3-color gradient scale across `range`'s values — e.g. "color-scale
 * the Total Amount column". No `format` here — each stop supplies its own
 * color directly. The lowest/highest actual values anchor the scale's ends
 * and shift automatically as data changes (TASKS.md #37).
 */
export interface ConditionalFormatColorScaleRule {
  kind: 'colorScale';
  /** [min, max] for a 2-color scale, or [min, mid, max] for a 3-color scale. */
  colors: [string, string] | [string, string, string];
}

export type ConditionalFormatRule =
  | ConditionalFormatCellValueRule
  | ConditionalFormatFormulaRule
  | ConditionalFormatTopBottomRule
  | ConditionalFormatColorScaleRule;

export type RangeFilterOperator =
  | 'equals'
  | 'contains'
  | 'greaterThan'
  | 'lessThan'
  | 'notEquals'
  | 'lengthEquals'
  | 'lengthNotEquals'
  | 'matchesRegex'
  | 'notMatchesRegex';

export interface RangeFilterSpec {
  column: string;
  operator: RangeFilterOperator;
  value: string | number;
}

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
  position?: 'above' | 'below' | 'left' | 'right' | 'before' | 'after' | 'afterLastColumn';
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
  /** INSERT_COLUMN semantic: insert after this header name */
  afterColumn?: string;
  columns?: string[];
  oldName?: string;
  sourceName?: string;
  newName?: string;
  name?: string;
  tableName?: string;
  hasHeaders?: boolean;
  style?: string;
  sourceSheetName?: string;
  chartType?: string;
  title?: string;
  startCell?: string;
  endCell?: string;
  question?: string;
  options?: string[];
  message?: string;
  copyFrom?: string;
  /** SORT_RANGE — 0-based column index within range */
  key?: number;
  ascending?: boolean;
  columnName?: string;
  /** SET_MATCHING_ROWS — column to write into */
  targetColumn?: string;
  /** COPY_FILTERED_RANGE / MOVE_RANGE / FORMAT_MATCHING_ROWS / SET_MATCHING_ROWS */
  sourceSheet?: string;
  destSheet?: string;
  destStartCell?: string;
  filter?: RangeFilterSpec;
  mode?: 'copy' | 'move';
  /** AGGREGATE_TABLE */
  groupByColumn?: string;
  groupByTransform?: 'none' | 'month' | 'year' | 'monthYear' | 'weekday' | 'quarter';
  aggregations?: Array<{
    column: string;
    /** 'first' passes through a label column 1:1 with the group key (e.g. Supplier Name alongside a GSTIN group-by) — not a numeric reduction. */
    fn: 'sum' | 'count' | 'average' | 'max' | 'min' | 'first';
    outputLabel: string;
  }>;
  sortBy?: { column: string; direction: 'asc' | 'desc' };
  topN?: number;
  /** CREATE_CHART / UPDATE_CHART */
  destCell?: string;
  colorScheme?: 'default' | 'blue' | 'grey' | 'blueGrey' | 'green' | 'red' | 'orange' | 'purple' | 'yellow';
  chartId?: string;
  /** CONDITIONAL_FORMAT — live cell-value comparison rule (TASKS.md M3). */
  rule?: ConditionalFormatRule;
  /** CONDITIONAL_FORMAT — targets an existing rule by id to modify in place rather than create (TASKS.md #38). */
  existingRuleId?: string;
  /** DELETE_CONDITIONAL_FORMAT — the real Excel-assigned rule id to delete (revert-only, TASKS.md #40). */
  ruleId?: string;
  /**
   * When true, allow writing over non-empty cells.
   * Only for unambiguously replace/clear intents — never inferred by the Executor.
   */
  explicitOverwriteConfirmed?: boolean;
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
