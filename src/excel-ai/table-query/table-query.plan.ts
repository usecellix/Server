import { resolveColumn, TableColumn, TableModel } from './table-model';
import {
  TABLE_FILTER_OPS,
  TABLE_GROUP_BYS,
  TABLE_QUERY_OPS,
  TableFilter,
  TableFilterOp,
  TableGroupBy,
  TableQuery,
  TableQueryOp,
  TableQueryPlanError,
} from './table-query.types';

/**
 * Turns the model's JSON into a plan that is safe to run: every column must be
 * a real column of this sheet and every operation one the executor knows. A
 * plan that fails here is discarded, never repaired by guessing.
 */

const OP_ALIASES: Record<string, TableQueryOp> = {
  total: 'sum', add: 'sum', avg: 'average', mean: 'average', minimum: 'min', lowest: 'min', smallest: 'min',
  earliest: 'min', maximum: 'max', highest: 'max', largest: 'max', latest: 'max', biggest: 'max',
  top_n: 'top', bottom_n: 'bottom', rows: 'list', filter: 'list', unique: 'distinct',
};

const FILTER_ALIASES: Record<string, TableFilterOp> = {
  '=': 'equals', '==': 'equals', eq: 'equals', is: 'equals', '!=': 'not_equals', ne: 'not_equals', neq: 'not_equals',
  '>': 'gt', '>=': 'gte', '<': 'lt', '<=': 'lte', includes: 'contains', like: 'contains',
  excludes: 'not_contains', startswith: 'starts_with', endswith: 'ends_with', empty: 'blank',
  is_blank: 'blank', not_empty: 'not_blank', is_not_blank: 'not_blank', notblank: 'not_blank', range: 'between',
};

const MAX_QUERIES = 6;
const MAX_LIMIT = 50;

const key = (value: unknown): string => String(value ?? '').trim().toLowerCase().replace(/[\s-]+/g, '_');

function column(table: TableModel, name: unknown, role: string): TableColumn {
  const found = resolveColumn(table, name);
  if (!found) throw new TableQueryPlanError(`${role} column "${String(name ?? '')}" is not on the sheet`);
  return found;
}

function parseFilter(table: TableModel, raw: unknown): TableFilter {
  const f = (raw ?? {}) as Record<string, unknown>;
  const opKey = key(f.op ?? f.operator);
  const op = (TABLE_FILTER_OPS as readonly string[]).includes(opKey) ? (opKey as TableFilterOp) : FILTER_ALIASES[opKey];
  if (!op) throw new TableQueryPlanError(`unknown filter "${String(f.op ?? f.operator)}"`);
  const filter: TableFilter = { column: column(table, f.column, 'Filter'), op };
  const scalar = (v: unknown) => (typeof v === 'number' || typeof v === 'string' ? v : undefined);
  if (op !== 'blank' && op !== 'not_blank') {
    filter.value = scalar(f.value);
    if (filter.value === undefined || filter.value === '') throw new TableQueryPlanError(`filter "${op}" has no value`);
  }
  if (op === 'between') {
    filter.value2 = scalar(f.value2 ?? f.to ?? f.max);
    if (filter.value2 === undefined || filter.value2 === '') throw new TableQueryPlanError('filter "between" needs two values');
  }
  return filter;
}

function parseQuery(table: TableModel, raw: unknown): TableQuery {
  const q = (raw ?? {}) as Record<string, unknown>;
  const opKey = key(q.op ?? q.operation);
  const op = (TABLE_QUERY_OPS as readonly string[]).includes(opKey) ? (opKey as TableQueryOp) : OP_ALIASES[opKey];
  if (!op) throw new TableQueryPlanError(`unknown operation "${String(q.op ?? q.operation)}"`);

  const query: TableQuery = { op, filters: [] };
  const needsColumn = op !== 'count' && op !== 'list';
  if (needsColumn || (q.column !== undefined && q.column !== null && q.column !== '')) {
    query.column = column(table, q.column, 'Measure');
  }

  const filters = q.filters ?? q.where;
  if (filters !== undefined && filters !== null) {
    if (!Array.isArray(filters)) throw new TableQueryPlanError('filters must be a list');
    query.filters = filters.map((f) => parseFilter(table, f));
  }

  const group = (q.groupBy ?? q.group_by) as Record<string, unknown> | string | undefined | null;
  if (group) {
    const groupColumn = column(table, typeof group === 'string' ? group : group.column, 'Group');
    const byKey = typeof group === 'string' ? '' : key(group.by ?? group.transform);
    let by: TableGroupBy = (TABLE_GROUP_BYS as readonly string[]).includes(byKey) ? (byKey as TableGroupBy) : 'value';
    if (groupColumn.type !== 'date') by = 'value';
    // Summing a column "for each" of its own values is never what a question
    // means ("total debits per month" came back as one line per Debit amount).
    if (query.column && query.column.index === groupColumn.index && query.op !== 'count') {
      throw new TableQueryPlanError(`"${groupColumn.name}" is both the measure and the grouping`);
    }
    query.groupBy = { column: groupColumn, by };
  }

  if (q.limit !== undefined && q.limit !== null) {
    const limit = Number(q.limit);
    if (!Number.isInteger(limit) || limit < 1) throw new TableQueryPlanError(`limit "${String(q.limit)}" is not a positive whole number`);
    query.limit = Math.min(limit, MAX_LIMIT);
  }
  return query;
}

export type ParsedPlan = { kind: 'queries'; queries: TableQuery[] } | { kind: 'unsupported'; reason: string };

/** Reads the model's reply. Throws TableQueryPlanError when it is not a usable plan. */
export function parseTableQueryPlan(table: TableModel, reply: string): ParsedPlan {
  const start = reply.indexOf('{');
  const end = reply.lastIndexOf('}');
  if (start < 0 || end <= start) throw new TableQueryPlanError('reply is not JSON');
  let raw: Record<string, unknown>;
  try {
    raw = JSON.parse(reply.slice(start, end + 1)) as Record<string, unknown>;
  } catch {
    throw new TableQueryPlanError('reply is not valid JSON');
  }
  if (typeof raw.unsupported === 'string' && raw.unsupported.trim()) {
    return { kind: 'unsupported', reason: raw.unsupported.trim() };
  }
  const list = Array.isArray(raw.queries) ? raw.queries : raw.op ? [raw] : null;
  if (!list || list.length === 0) throw new TableQueryPlanError('plan has no queries');
  if (list.length > MAX_QUERIES) throw new TableQueryPlanError(`plan has ${list.length} queries, the most allowed is ${MAX_QUERIES}`);
  return { kind: 'queries', queries: list.map((q) => parseQuery(table, q)) };
}
