export type CellValue = string | number | boolean | null;

export interface FormatSpec {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  numberFormat?: string;
  horizontalAlignment?: 'left' | 'center' | 'right';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  borders?: {
    style: 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'none';
    color?: string;
    edges: ('top' | 'bottom' | 'left' | 'right' | 'all' | 'outer' | 'inner')[];
  };
}

export interface BatchSetOperation {
  address: string;
  value?: CellValue;
  formula?: string;
  format?: FormatSpec;
}

export interface BatchSetAction {
  type: 'BATCH_SET';
  sheetName: string;
  operations: BatchSetOperation[];
}

export interface CreateTableAction {
  type: 'CREATE_TABLE';
  sheetName: string;
  range: string;
  tableName: string;
  hasHeaders: boolean;
  style?: string;
}

export interface DefineNamedRangeAction {
  type: 'DEFINE_NAMED_RANGE';
  name: string;
  formula: string;
  comment?: string;
}

export interface AutoFitColumnsAction {
  type: 'AUTOFIT_COLUMNS';
  sheetName: string;
  columns?: string[];
}

export interface ClarifyAction {
  type: 'CLARIFY';
  question: string;
  options?: string[];
}

export interface CheckpointAction {
  type: 'CHECKPOINT';
  message: string;
}

export type RichStructuralAction =
  | BatchSetAction
  | CreateTableAction
  | DefineNamedRangeAction
  | AutoFitColumnsAction
  | ClarifyAction
  | CheckpointAction;
