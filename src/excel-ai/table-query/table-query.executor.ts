import { isBlank, TableColumn, TableModel, TableRow, toIsoDateValue, toNumber } from './table-model';
import {
  GroupsResult,
  TableFilter,
  TableGroupBy,
  TableQuery,
  TableQueryPlanError,
  TableQueryResult,
} from './table-query.types';

/**
 * Runs a query plan over every row of a table. Pure and deterministic: the
 * same plan on the same sheet always gives the same answer.
 */

const DEFAULT_TOP = 5;
const DEFAULT_LIST = 20;
const DEFAULT_GROUPS = 24;
const MAX_ROWS_RETURNED = 50;

const textOf = (value: unknown): string => (isBlank(value) ? '' : String(value).trim());

/** The cell as this column's own type, or null when it is blank or does not fit. */
function typed(column: TableColumn, row: TableRow): number | string | null {
  const raw = row.cells[column.index];
  if (isBlank(raw)) return null;
  if (column.type === 'number') return toNumber(raw);
  if (column.type === 'date') return toIsoDateValue(raw);
  return textOf(raw);
}

/** Adds decimals without binary drift: 0.1 + 0.2 is 0.3 here. */
export function sumExact(values: number[]): number {
  let scale = 1;
  for (const value of values) {
    const text = String(value);
    if (text.includes('e')) return values.reduce((a, b) => a + b, 0);
    const places = text.split('.')[1]?.length ?? 0;
    scale = Math.max(scale, 10 ** Math.min(places, 6));
  }
  let total = 0;
  for (const value of values) total += Math.round(value * scale);
  return total / scale;
}

function comparable(column: TableColumn, value: string | number | undefined): number | string | null {
  if (value === undefined || value === null || value === '') return null;
  if (column.type === 'number') return toNumber(value);
  if (column.type === 'date') return toIsoDateValue(value);
  return String(value).trim().toLowerCase();
}

function matches(filter: TableFilter, row: TableRow): boolean {
  const { column, op } = filter;
  const raw = row.cells[column.index];
  if (op === 'blank') return isBlank(raw);
  if (op === 'not_blank') return !isBlank(raw);

  const text = textOf(raw).toLowerCase();
  const needle = String(filter.value ?? '').trim().toLowerCase();
  switch (op) {
    case 'contains':
      return needle !== '' && text.includes(needle);
    case 'not_contains':
      return needle === '' || !text.includes(needle);
    case 'starts_with':
      return needle !== '' && text.startsWith(needle);
    case 'ends_with':
      return needle !== '' && text.endsWith(needle);
    default:
      break;
  }

  const cell = column.type === 'text' ? text : typed(column, row);
  const a = comparable(column, filter.value);
  if (op === 'equals') return a !== null && cell !== null && cell === a;
  if (op === 'not_equals') return a === null || cell === null || cell !== a;
  if (cell === null || a === null) return false;
  if (op === 'gt') return cell > a;
  if (op === 'gte') return cell >= a;
  if (op === 'lt') return cell < a;
  if (op === 'lte') return cell <= a;
  if (op === 'between') {
    const b = comparable(column, filter.value2);
    return b !== null && cell >= a && cell <= b;
  }
  return false;
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function groupKey(column: TableColumn, by: TableGroupBy, row: TableRow): { key: string; label: string } | null {
  const value = typed(column, row);
  if (value === null) return null;
  if (column.type === 'date' && by !== 'value') {
    const iso = String(value);
    if (by === 'year') return { key: iso.slice(0, 4), label: iso.slice(0, 4) };
    if (by === 'month') {
      return { key: iso.slice(0, 7), label: `${MONTHS[Number(iso.slice(5, 7)) - 1]} ${iso.slice(0, 4)}` };
    }
    return { key: iso, label: iso };
  }
  const label = String(value);
  return { key: column.type === 'text' ? label.toLowerCase() : label, label };
}

function aggregate(op: TableQuery['op'], column: TableColumn | undefined, rows: TableRow[]): number | null {
  if (op === 'count') {
    return column ? rows.filter((row) => !isBlank(row.cells[column.index])).length : rows.length;
  }
  if (!column) return null;
  const numbers = rows.map((row) => typed(column, row)).filter((v): v is number => typeof v === 'number');
  if (!numbers.length) return null;
  if (op === 'sum') return sumExact(numbers);
  if (op === 'average') return sumExact(numbers) / numbers.length;
  if (op === 'min') return Math.min(...numbers);
  if (op === 'max') return Math.max(...numbers);
  return null;
}

function runGrouped(query: TableQuery, rows: TableRow[]): GroupsResult {
  const { column: groupColumn, by } = query.groupBy as NonNullable<TableQuery['groupBy']>;
  if (!['sum', 'average', 'count', 'min', 'max'].includes(query.op)) {
    throw new TableQueryPlanError(`"${query.op}" cannot be grouped`);
  }
  if (query.op !== 'count' && query.column?.type !== 'number') {
    throw new TableQueryPlanError(`"${query.op}" needs a number column`);
  }
  const buckets = new Map<string, { label: string; rows: TableRow[] }>();
  for (const row of rows) {
    const group = groupKey(groupColumn, by, row);
    if (!group) continue;
    const bucket = buckets.get(group.key) ?? { label: group.label, rows: [] };
    bucket.rows.push(row);
    buckets.set(group.key, bucket);
  }
  let groups = [...buckets.entries()].map(([key, bucket]) => ({
    key,
    label: bucket.label,
    value: aggregate(query.op, query.column, bucket.rows),
    rows: bucket.rows.length,
  }));
  // Dates read in order of time; anything else reads largest first.
  if (groupColumn.type === 'date') groups.sort((a, b) => a.key.localeCompare(b.key));
  else groups.sort((a, b) => (b.value ?? -Infinity) - (a.value ?? -Infinity) || a.label.localeCompare(b.label));

  const limit = Math.min(query.limit ?? DEFAULT_GROUPS, MAX_ROWS_RETURNED);
  const groupsOmitted = Math.max(0, groups.length - limit);
  groups = groups.slice(0, limit);
  return { kind: 'groups', query, groups, matchedRows: rows.length, groupsOmitted };
}

export function executeTableQuery(table: TableModel, query: TableQuery): TableQueryResult {
  const rows = table.rows.filter((row) => query.filters.every((filter) => matches(filter, row)));
  const { op, column } = query;

  if (query.groupBy) return runGrouped(query, rows);

  if (op === 'count') {
    return { kind: 'scalar', query, value: aggregate('count', column, rows), valueType: 'number', matchedRows: rows.length };
  }

  if (op === 'list') {
    const limit = Math.min(query.limit ?? DEFAULT_LIST, MAX_ROWS_RETURNED);
    return { kind: 'rows', query, rows: rows.slice(0, limit), matchedRows: rows.length };
  }

  if (!column) throw new TableQueryPlanError(`"${op}" needs a column`);

  if (op === 'distinct') {
    const counts = new Map<string, { label: string; rows: number }>();
    for (const row of rows) {
      const value = typed(column, row);
      if (value === null) continue;
      const label = String(value);
      const key = column.type === 'text' ? label.toLowerCase() : label;
      const entry = counts.get(key) ?? { label, rows: 0 };
      entry.rows += 1;
      counts.set(key, entry);
    }
    const values = [...counts.values()].sort((a, b) => b.rows - a.rows || a.label.localeCompare(b.label));
    const limit = Math.min(query.limit ?? DEFAULT_LIST, MAX_ROWS_RETURNED);
    return { kind: 'values', query, values: values.slice(0, limit), distinctCount: values.length, matchedRows: rows.length };
  }

  if (op === 'first' || op === 'last') {
    const present = rows.filter((row) => typed(column, row) !== null);
    const row = op === 'first' ? present[0] : present[present.length - 1];
    return { kind: 'scalar', query, value: row ? typed(column, row) : null, valueType: column.type, matchedRows: rows.length, row };
  }

  if (op === 'sum' || op === 'average') {
    if (column.type !== 'number') throw new TableQueryPlanError(`"${op}" needs a number column, and ${column.name} is ${column.type}`);
    const valuesUsed = rows.filter((row) => typeof typed(column, row) === 'number').length;
    return { kind: 'scalar', query, value: aggregate(op, column, rows), valueType: 'number', matchedRows: rows.length, valuesUsed };
  }

  // min, max, top, bottom: rank the rows that have a value in the column.
  if (column.type === 'text') throw new TableQueryPlanError(`"${op}" needs a number or date column, and ${column.name} is text`);
  const ranked = rows
    .map((row) => ({ row, value: typed(column, row) }))
    .filter((entry): entry is { row: TableRow; value: number | string } => entry.value !== null);
  const descending = op === 'max' || op === 'top';
  ranked.sort((a, b) => {
    const order = a.value < b.value ? -1 : a.value > b.value ? 1 : 0;
    // Equal values keep sheet order, so the same question always lists the same rows.
    return (descending ? -order : order) || a.row.rowNumber - b.row.rowNumber;
  });

  if (op === 'min' || op === 'max') {
    const best = ranked[0];
    return {
      kind: 'scalar',
      query,
      value: best ? best.value : null,
      valueType: column.type,
      matchedRows: rows.length,
      row: best?.row,
      tiedRows: best ? ranked.filter((entry) => entry.value === best.value).length - 1 : 0,
    };
  }

  const limit = Math.min(query.limit ?? DEFAULT_TOP, MAX_ROWS_RETURNED);
  const kept = ranked.slice(0, limit);
  const lastValue = kept[kept.length - 1]?.value;
  const tiedBeyondLimit = lastValue === undefined ? 0 : ranked.slice(limit).filter((entry) => entry.value === lastValue).length;
  return { kind: 'rows', query, rows: kept.map((entry) => entry.row), matchedRows: rows.length, tiedBeyondLimit };
}
