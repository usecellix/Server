import { isBlank, TableColumn, TableModel, TableRow, toIsoDateValue, toNumber } from './table-model';
import { TableFilter, TableQuery, TableQueryResult } from './table-query.types';

/**
 * Writes the answer from computed results. Every figure in the text is printed
 * here from the value the executor returned, so nothing can be mis-copied, and
 * each result says in words what was computed, so a wrong reading of the
 * question is visible instead of hidden.
 */

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

export function formatDate(iso: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  return match ? `${match[3]}-${MONTHS[Number(match[2]) - 1]}-${match[1]}` : iso;
}

/** Indian digit grouping (12,34,567.89), which is what the sheets this product serves use. */
export function formatNumber(value: number, decimals: number): string {
  return new Intl.NumberFormat('en-IN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
}

const clip = (text: string, max: number) => (text.length > max ? `${text.slice(0, max - 1).trimEnd()}…` : text);

function formatValue(column: TableColumn, raw: unknown): string {
  if (isBlank(raw)) return '';
  if (column.type === 'number') {
    const n = toNumber(raw);
    return n === null ? String(raw) : formatNumber(n, column.decimals);
  }
  if (column.type === 'date') {
    const iso = toIsoDateValue(raw);
    return iso ? formatDate(iso) : String(raw);
  }
  return String(raw).trim();
}

function formatFilterValue(column: TableColumn, value: string | number | undefined): string {
  if (value === undefined) return '';
  if (column.type === 'text') return `"${value}"`;
  return formatValue(column, value) || String(value);
}

const FILTER_WORDS: Record<TableFilter['op'], string> = {
  contains: 'contains', not_contains: 'does not contain', equals: 'is', not_equals: 'is not', starts_with: 'starts with',
  ends_with: 'ends with', gt: 'is more than', gte: 'is at least', lt: 'is less than', lte: 'is at most', between: 'is between',
  blank: 'is blank', not_blank: 'is not blank',
};

function describeFilters(filters: TableFilter[]): string {
  if (!filters.length) return '';
  const parts = filters.map((f) => {
    if (f.op === 'blank' || f.op === 'not_blank') return `${f.column.name} ${FILTER_WORDS[f.op]}`;
    if (f.op === 'between') return `${f.column.name} is between ${formatFilterValue(f.column, f.value)} and ${formatFilterValue(f.column, f.value2)}`;
    return `${f.column.name} ${FILTER_WORDS[f.op]} ${formatFilterValue(f.column, f.value)}`;
  });
  return ` where ${parts.join(' and ')}`;
}

function describeMeasure(query: TableQuery, count?: number): string {
  const name = query.column?.name ?? '';
  const isDate = query.column?.type === 'date';
  switch (query.op) {
    case 'sum': return `Total of ${name}`;
    case 'average': return `Average ${name}`;
    case 'count': {
      // "with a Debit" says something. "with a Description where Description contains…" does not.
      const adds = query.column && query.column.type !== 'text' && !query.filters.some((f) => f.column === query.column);
      return adds ? `Number of rows with a ${name}` : 'Number of rows';
    }
    case 'min': return isDate ? `Earliest ${name}` : `Lowest ${name}`;
    case 'max': return isDate ? `Latest ${name}` : `Highest ${name}`;
    case 'first': return `First ${name}`;
    case 'last': return `Last ${name}`;
    case 'top': return isDate ? `${count} latest by ${name}` : `${count} largest by ${name}`;
    case 'bottom': return isDate ? `${count} earliest by ${name}` : `${count} smallest by ${name}`;
    case 'distinct': return `Different values of ${name}`;
    default: return 'Matching rows';
  }
}

function describeGroup(query: TableQuery): string {
  if (!query.groupBy) return '';
  const { column, by } = query.groupBy;
  return by === 'value' ? ` for each ${column.name}` : ` for each ${by} of ${column.name}`;
}

const rowsWord = (n: number) => `${formatNumber(n, 0)} ${n === 1 ? 'row' : 'rows'}`;

/** The columns worth showing for a listed row: when, what, and the figure asked about. */
function displayColumns(table: TableModel, query: TableQuery): TableColumn[] {
  const picked: TableColumn[] = [];
  const add = (c: TableColumn | undefined) => {
    if (c && !picked.includes(c)) picked.push(c);
  };
  add(table.columns.find((c) => c.type === 'date'));
  // The description is the text column with the longest values.
  const sample = table.rows.slice(0, 60);
  let best: { column: TableColumn; length: number } | null = null;
  for (const c of table.columns.filter((col) => col.type === 'text')) {
    const lengths = sample.map((row) => String(row.cells[c.index] ?? '').trim().length).filter((n) => n > 0);
    if (!lengths.length) continue;
    const average = lengths.reduce((a, b) => a + b, 0) / lengths.length;
    if (average >= 6 && (!best || average > best.length)) best = { column: c, length: average };
  }
  add(best?.column);
  add(query.column);
  for (const f of query.filters) if (picked.length < 4) add(f.column);
  if (picked.length < 2) for (const c of table.columns) if (picked.length < 3) add(c);
  return picked.slice(0, 4);
}

function describeRow(table: TableModel, query: TableQuery, row: TableRow, skip?: TableColumn): string {
  const parts: string[] = [];
  for (const c of displayColumns(table, query)) {
    if (c === skip) continue;
    const text = formatValue(c, row.cells[c.index]);
    if (!text) continue;
    parts.push(c.type === 'number' ? `${c.name} ${text}` : clip(text, 70));
  }
  return `${parts.join(' · ')} (row ${row.rowNumber})`;
}

function formatResult(table: TableModel, result: TableQueryResult): string[] {
  const { query } = result;
  const where = describeFilters(query.filters);

  if (result.kind === 'scalar') {
    const label = `${describeMeasure(query)}${where}`;
    if (result.value === null) {
      return [`**${label}:** nothing to work from. ${result.matchedRows === 0 ? 'No rows match.' : 'The matching rows have no value in that column.'}`];
    }
    const column = query.column;
    const text =
      result.valueType === 'number'
        ? formatNumber(result.value as number, query.op === 'count' ? 0 : column?.decimals ?? 2)
        : result.valueType === 'date'
          ? formatDate(String(result.value))
          : String(result.value);
    const lines = [`**${label}: ${text}**`];
    // An earliest or latest date is the answer in itself; which row holds it adds nothing.
    const isDateExtreme = column?.type === 'date' && (query.op === 'min' || query.op === 'max');
    if (result.row && !isDateExtreme) {
      const tie = result.tiedRows ? ` ${rowsWord(result.tiedRows)} more ${result.tiedRows === 1 ? 'has' : 'have'} the same value.` : '';
      lines.push(`   ${describeRow(table, query, result.row, column)}.${tie}`);
    } else if (query.op === 'sum' || query.op === 'average') {
      lines.push(`   From ${rowsWord(result.valuesUsed ?? result.matchedRows)}.`);
    }
    return lines;
  }

  if (result.kind === 'rows') {
    const shown = result.rows.length;
    const label = `${describeMeasure(query, shown)}${where}`;
    if (shown === 0) return [`**${label}:** no rows match.`];
    const lines = [`**${label}**`];
    result.rows.forEach((row, i) => lines.push(`${i + 1}. ${describeRow(table, query, row)}`));
    if (result.tiedBeyondLimit) {
      const last = result.rows[shown - 1];
      const value = query.column ? formatValue(query.column, last.cells[query.column.index]) : '';
      lines.push(`${rowsWord(result.tiedBeyondLimit)} more ${result.tiedBeyondLimit === 1 ? 'has' : 'have'} the same ${query.column?.name ?? 'value'} (${value}) as the last one listed.`);
    }
    if (query.op === 'list' && result.matchedRows > shown) {
      lines.push(`Showing the first ${shown} of ${rowsWord(result.matchedRows)} that match.`);
    }
    return lines;
  }

  if (result.kind === 'groups') {
    const label = `${describeMeasure(query)}${describeGroup(query)}${where}`;
    if (!result.groups.length) return [`**${label}:** no rows match.`];
    const decimals = query.op === 'count' ? 0 : query.column?.decimals ?? 2;
    const lines = [`**${label}**`];
    for (const g of result.groups) {
      const name = query.groupBy?.column.type === 'date' && query.groupBy.by === 'day' ? formatDate(g.label) : clip(g.label, 60);
      lines.push(`- ${name}: ${g.value === null ? 'no values' : formatNumber(g.value, decimals)}${query.op === 'count' ? '' : ` (${rowsWord(g.rows)})`}`);
    }
    if (result.groupsOmitted) lines.push(`${result.groupsOmitted} more not shown.`);
    return lines;
  }

  const label = `${describeMeasure(query)}${where}`;
  const lines = [`**${label}: ${formatNumber(result.distinctCount, 0)}**`];
  for (const v of result.values) lines.push(`- ${clip(v.label, 60)} (${rowsWord(v.rows)})`);
  if (result.distinctCount > result.values.length) lines.push(`${result.distinctCount - result.values.length} more not shown.`);
  return lines;
}

/**
 * The sheet rows an answer names with "(row N)", in the order it names them and
 * without repeats. The add-in turns each into a pointer that selects the row.
 * Kept in step with `formatResult`: a row is listed here exactly when its
 * "(row N)" is printed there.
 */
export function rowsMentioned(results: TableQueryResult[]): number[] {
  const seen = new Set<number>();
  for (const result of results) {
    if (result.kind === 'rows') {
      for (const row of result.rows) seen.add(row.rowNumber);
    } else if (result.kind === 'scalar' && result.row && result.value !== null) {
      const { op, column } = result.query;
      const isDateExtreme = column?.type === 'date' && (op === 'min' || op === 'max');
      if (!isDateExtreme) seen.add(result.row.rowNumber);
    }
  }
  return [...seen];
}

export interface FormatOptions {
  /** Rows the sheet says it has, header excluded. Used to say so when not all could be read. */
  expectedRows?: number;
}

export function formatTableAnswer(table: TableModel, results: TableQueryResult[], options: FormatOptions = {}): string {
  const blocks = results.map((result) => formatResult(table, result).join('\n'));
  const read = table.rows.length;
  const expected = options.expectedRows ?? read;
  const scope =
    expected > read
      ? `Only ${rowsWord(read)} of ${formatNumber(expected, 0)} on ${table.sheetName} could be read, so this may be incomplete.`
      : `Worked out from all ${rowsWord(read)} of ${table.sheetName}.`;
  return `${blocks.join('\n\n')}\n\n${scope}`;
}
