import { TableColumn, TableRow } from './table-model';

/**
 * What a question about a table is translated into. The model writes this
 * plan; code runs it. Nothing here carries a computed value.
 */

export const TABLE_QUERY_OPS = [
  'sum', 'average', 'count', 'min', 'max', 'top', 'bottom', 'list', 'first', 'last', 'distinct',
] as const;
export type TableQueryOp = (typeof TABLE_QUERY_OPS)[number];

export const TABLE_FILTER_OPS = [
  'contains', 'not_contains', 'equals', 'not_equals', 'starts_with', 'ends_with',
  'gt', 'gte', 'lt', 'lte', 'between', 'blank', 'not_blank',
] as const;
export type TableFilterOp = (typeof TABLE_FILTER_OPS)[number];

export const TABLE_GROUP_BYS = ['value', 'day', 'month', 'year'] as const;
export type TableGroupBy = (typeof TABLE_GROUP_BYS)[number];

export interface TableFilter {
  column: TableColumn;
  op: TableFilterOp;
  value?: string | number;
  value2?: string | number;
}

export interface TableQuery {
  op: TableQueryOp;
  /** The column measured, ranked or read. Optional only for `count` and `list`. */
  column?: TableColumn;
  filters: TableFilter[];
  groupBy?: { column: TableColumn; by: TableGroupBy };
  limit?: number;
}

export interface ScalarResult {
  kind: 'scalar';
  query: TableQuery;
  /** A number, a yyyy-mm-dd date, text, or null when nothing matched. */
  value: number | string | null;
  valueType: 'number' | 'date' | 'text';
  matchedRows: number;
  /** For sum and average: how many rows had a number in the column. */
  valuesUsed?: number;
  /** The row the value came from, for min, max, first and last. */
  row?: TableRow;
  /** Other rows holding the same extreme value. */
  tiedRows?: number;
}

export interface RowsResult {
  kind: 'rows';
  query: TableQuery;
  rows: TableRow[];
  matchedRows: number;
  /** Rows left out of a top or bottom list that equal its last value. */
  tiedBeyondLimit?: number;
}

export interface GroupsResult {
  kind: 'groups';
  query: TableQuery;
  groups: Array<{ key: string; label: string; value: number | null; rows: number }>;
  matchedRows: number;
  /** Groups not shown because of the limit. */
  groupsOmitted: number;
}

export interface ValuesResult {
  kind: 'values';
  query: TableQuery;
  values: Array<{ label: string; rows: number }>;
  distinctCount: number;
  matchedRows: number;
}

export type TableQueryResult = ScalarResult | RowsResult | GroupsResult | ValuesResult;

export class TableQueryPlanError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'TableQueryPlanError';
  }
}
