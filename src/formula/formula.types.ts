export type FormulaFunctionName =
  | 'SUM'
  | 'SUMIF'
  | 'SUMIFS'
  | 'AVERAGE'
  | 'AVERAGEIF'
  | 'AVERAGEIFS'
  | 'COUNT'
  | 'COUNTA'
  | 'COUNTIF'
  | 'COUNTIFS'
  | 'IF'
  | 'IFS'
  | 'IFERROR'
  | 'IFNA'
  | 'AND'
  | 'OR'
  | 'NOT'
  | 'VLOOKUP'
  | 'HLOOKUP'
  | 'INDEX'
  | 'MATCH'
  | 'XLOOKUP'
  | 'MAX'
  | 'MIN'
  | 'LARGE'
  | 'SMALL'
  | 'ROUND'
  | 'ROUNDUP'
  | 'ROUNDDOWN'
  | 'NPV'
  | 'IRR'
  | 'PMT'
  | 'FV'
  | 'PV'
  | 'TEXT'
  | 'LEFT'
  | 'RIGHT'
  | 'MID'
  | 'CONCAT'
  | 'LEN'
  | 'DATE'
  | 'YEAR'
  | 'MONTH'
  | 'DAY'
  | 'EDATE'
  | 'EOMONTH';

export interface CellRef {
  sheet?: string;
  column: string;
  row: number;
  isAbsoluteCol: boolean;
  isAbsoluteRow: boolean;
  raw: string;
}

export interface RangeRef {
  sheet?: string;
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
  isAbsoluteStart: boolean;
  isAbsoluteEnd: boolean;
  raw: string;
}

export interface ParsedFormula {
  raw: string;
  functions: FormulaFunctionName[];
  cellRefs: CellRef[];
  rangeRefs: RangeRef[];
  crossSheetRefs: string[];
  isAggregation: boolean;
  isLookup: boolean;
  isConditional: boolean;
}

export interface FormulaNode {
  address: string;
  sheetName: string;
  formula: ParsedFormula;
  rowIndex: number;
  colIndex: number;
}

export interface RowPattern {
  rowIndex: number;
  type: 'data' | 'total' | 'subtotal' | 'header' | 'empty' | 'label';
  formulaTypes: FormulaFunctionName[];
  description: string;
}

export interface ColumnPattern {
  colLetter: string;
  type: 'data' | 'formula' | 'label' | 'mixed';
  dominantFunction?: FormulaFunctionName;
  description: string;
}

export interface FormulaInsights {
  sheetName: string;
  totalFormulas: number;
  crossSheetRefs: string[];
  functionsSummary: Record<string, number>;
  rowPatterns: RowPattern[];
  columnPatterns: ColumnPattern[];
  aggregationRows: number[];
  dataRowRange?: { start: number; end: number };
  dependencyWarnings: string[];
  llmSummary: string;
}

export type FormulaValidationCode =
  | 'SYNTAX'
  | 'REFERENCE'
  | 'NAMED_RANGE'
  | 'POST_EXEC';

export interface FormulaValidationIssue {
  severity: 'error' | 'warning';
  code: FormulaValidationCode;
  actionIndex?: number;
  formula?: string;
  cell?: string;
  message: string;
  suggestion?: string;
}

export interface FormulaValidationResult {
  passed: boolean;
  issues: FormulaValidationIssue[];
  phase: 'pre_apply' | 'post_apply';
}
