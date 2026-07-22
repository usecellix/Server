export type AggregateFn = 'sum' | 'count' | 'average' | 'max' | 'min';

/** Derived grouping key transforms (computed in code during aggregation). */
export type GroupByTransform = 'none' | 'month' | 'year' | 'monthYear' | 'weekday' | 'quarter';

export interface AggregateSpec {
  column: string;
  fn: AggregateFn;
  outputLabel: string;
}

export interface AggregateTableParams {
  rows: unknown[][];
  hasHeaders: boolean;
  groupByColumn: string;
  /** When set, group by a derived value of groupByColumn (e.g. month-of-year). */
  groupByTransform?: GroupByTransform;
  aggregations: AggregateSpec[];
  sortBy?: { column: string; direction: 'asc' | 'desc' };
  topN?: number;
}

function toNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim() !== '') {
    const n = Number(value.replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function findCol(headerRow: unknown[], name: string): number {
  const target = name.trim().toLowerCase();
  return headerRow.findIndex((cell) => String(cell ?? '').trim().toLowerCase() === target);
}

function aggregateValues(values: number[], fn: AggregateFn): number {
  if (fn === 'count') return values.length;
  if (values.length === 0) return 0;
  if (fn === 'sum') return values.reduce((a, b) => a + b, 0);
  if (fn === 'average') return values.reduce((a, b) => a + b, 0) / values.length;
  if (fn === 'max') return Math.max(...values);
  return Math.min(...values);
}

function parseDateLike(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number' && Number.isFinite(value)) {
    // Excel serial date (days since 1899-12-30) — only treat large numbers as serials.
    if (value > 20000 && value < 80000) {
      const epoch = Date.UTC(1899, 11, 30);
      return new Date(epoch + value * 86400000);
    }
  }
  if (typeof value === 'string' && value.trim()) {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

const MONTH_NAMES = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

const WEEKDAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

/**
 * Transform a raw cell value into a grouping key.
 * Office.js / backend compute this — the LLM only specifies the transform name.
 */
export function applyGroupByTransform(
  value: unknown,
  transform: GroupByTransform | undefined,
): string {
  const raw = String(value ?? '').trim();
  if (!transform || transform === 'none') return raw;

  const date = parseDateLike(value);
  if (!date) return raw;

  switch (transform) {
    case 'month':
      return MONTH_NAMES[date.getUTCMonth()] ?? raw;
    case 'year':
      return String(date.getUTCFullYear());
    case 'monthYear':
      return `${MONTH_NAMES[date.getUTCMonth()]} ${date.getUTCFullYear()}`;
    case 'weekday':
      return WEEKDAY_NAMES[date.getUTCDay()] ?? raw;
    case 'quarter':
      return `Q${Math.floor(date.getUTCMonth() / 3) + 1}`;
    default:
      return raw;
  }
}

function groupHeaderLabel(groupByColumn: string, transform?: GroupByTransform): string {
  if (!transform || transform === 'none') return groupByColumn;
  if (transform === 'month') return 'Month';
  if (transform === 'year') return 'Year';
  if (transform === 'monthYear') return 'Month';
  if (transform === 'weekday') return 'Weekday';
  if (transform === 'quarter') return 'Quarter';
  return groupByColumn;
}

/**
 * Group-by aggregate in memory — same pattern as COPY_FILTERED_RANGE (no LLM transcription).
 * Returns a 2D table: header row + data rows.
 */
export function buildAggregateTable(params: AggregateTableParams): unknown[][] {
  const { rows, hasHeaders, groupByColumn, groupByTransform, aggregations, sortBy, topN } = params;
  if (rows.length === 0) return [];

  const headerRow = hasHeaders ? rows[0] : null;
  const dataRows = hasHeaders ? rows.slice(1) : rows;
  if (!headerRow) {
    throw new Error('AGGREGATE_TABLE requires hasHeaders: true');
  }

  const groupIdx = findCol(headerRow, groupByColumn);
  if (groupIdx === -1) {
    throw new Error(`Group-by column "${groupByColumn}" not found`);
  }

  const aggCols = aggregations.map((agg) => {
    const idx = findCol(headerRow, agg.column);
    if (idx === -1) {
      throw new Error(`Aggregation column "${agg.column}" not found`);
    }
    return { ...agg, idx };
  });

  const groups = new Map<string, unknown[][]>();
  for (const row of dataRows) {
    const key = applyGroupByTransform(row[groupIdx], groupByTransform);
    if (!key) continue;
    const bucket = groups.get(key) ?? [];
    bucket.push(row);
    groups.set(key, bucket);
  }

  const outGroupHeader = groupHeaderLabel(groupByColumn, groupByTransform);
  const outHeader = [outGroupHeader, ...aggregations.map((a) => a.outputLabel)];
  let outRows: unknown[][] = [];

  for (const [key, bucket] of groups) {
    const cells: unknown[] = [key];
    for (const agg of aggCols) {
      const nums =
        agg.fn === 'count'
          ? bucket.map(() => 1)
          : bucket
              .map((row) => toNumber(row[agg.idx]))
              .filter((n): n is number => n !== null);
      cells.push(aggregateValues(nums, agg.fn));
    }
    outRows.push(cells);
  }

  if (sortBy) {
    const sortColIdx =
      sortBy.column.trim().toLowerCase() === groupByColumn.trim().toLowerCase() ||
      sortBy.column.trim().toLowerCase() === outGroupHeader.trim().toLowerCase()
        ? 0
        : outHeader.findIndex(
            (h) => String(h).trim().toLowerCase() === sortBy.column.trim().toLowerCase(),
          );
    if (sortColIdx >= 0) {
      const dir = sortBy.direction === 'desc' ? -1 : 1;
      outRows.sort((a, b) => {
        const av = a[sortColIdx];
        const bv = b[sortColIdx];
        const an = toNumber(av);
        const bn = toNumber(bv);
        if (an !== null && bn !== null) return (an - bn) * dir;
        return String(av ?? '').localeCompare(String(bv ?? '')) * dir;
      });
    }
  }

  if (typeof topN === 'number' && topN > 0) {
    outRows = outRows.slice(0, topN);
  }

  return [outHeader, ...outRows];
}
