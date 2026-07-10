export interface ShadowCell {
  value: unknown;
  formula: string;
  numberFormat: string;
}

export interface ShadowSheet {
  name: string;
  cells: Map<string, ShadowCell>;
  rowCount: number;
  columnCount: number;
  structure: string;
}

export interface ShadowWorkbook {
  activeSheetName: string;
  sheets: Map<string, ShadowSheet>;
  namedRanges: Map<string, string>;
  tables: string[];
  changedCells: Set<string>;
}
