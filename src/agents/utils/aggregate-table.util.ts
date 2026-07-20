export type AggregateFn = 'sum' | 'count' | 'average' | 'max' | 'min';

export interface AggregateSpec {
  column: string;
  fn: AggregateFn;
  outputLabel: string;
}

export interface AggregateTableParams {
  rows: unknown[][];
  hasHeaders: boolean;
  groupByColumn: string;
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

/**
 * Group-by aggregate in memory — same pattern as COPY_FILTERED_RANGE (no LLM transcription).
 * Returns a 2D table: header row + data rows.
 */
export function buildAggregateTable(params: AggregateTableParams): unknown[][] {
  const { rows, hasHeaders, groupByColumn, aggregations, sortBy, topN } = params;
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
    const key = String(row[groupIdx] ?? '').trim();
    if (!key) continue;
    const bucket = groups.get(key) ?? [];
    bucket.push(row);
    groups.set(key, bucket);
  }

  const outHeader = [groupByColumn, ...aggregations.map((a) => a.outputLabel)];
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
      sortBy.column.trim().toLowerCase() === groupByColumn.trim().toLowerCase()
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
