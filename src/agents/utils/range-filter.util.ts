import {
  RangeFilterOperator,
  RangeFilterSpec,
} from '../../excel-ai/types/sheet-actions.types';

export type { RangeFilterSpec };

export function applyFilterOperator(
  cellValue: unknown,
  filter: RangeFilterSpec,
): boolean {
  const { operator, value } = filter;

  if (operator === 'greaterThan' || operator === 'lessThan') {
    const left = Number(cellValue);
    const right = Number(value);
    if (!Number.isFinite(left) || !Number.isFinite(right)) return false;
    return operator === 'greaterThan' ? left > right : left < right;
  }

  if (operator === 'lengthEquals' || operator === 'lengthNotEquals') {
    const len = String(cellValue ?? '').trim().length;
    const expected = Number(value);
    if (!Number.isFinite(expected)) return false;
    return operator === 'lengthEquals' ? len === expected : len !== expected;
  }

  if (operator === 'matchesRegex' || operator === 'notMatchesRegex') {
    const text = String(cellValue ?? '').trim();
    try {
      const matched = new RegExp(String(value)).test(text);
      return operator === 'matchesRegex' ? matched : !matched;
    } catch {
      return false;
    }
  }

  const left = String(cellValue ?? '').trim().toLowerCase();
  const right = String(value ?? '').trim().toLowerCase();

  switch (operator as RangeFilterOperator) {
    case 'equals':
      return left === right;
    case 'notEquals':
      return left !== right;
    case 'contains':
      return left.includes(right);
    default:
      return false;
  }
}

export function resolveFilterColumnIndex(
  headerRow: unknown[],
  column: string | number,
): number {
  if (typeof column === 'number' && Number.isFinite(column)) {
    // Prefer 1-based Excel ordinals (K=11). Fall back to 0-based for index 0.
    if (Number.isInteger(column) && column >= 1 && column <= headerRow.length) {
      return column - 1;
    }
    if (Number.isInteger(column) && column >= 0 && column < headerRow.length) {
      return column;
    }
    throw new Error(`Column index ${column} is out of range for header row`);
  }

  const name = String(column ?? '').trim().toLowerCase();
  const colIndex = headerRow.findIndex(
    (cell) => String(cell ?? '').trim().toLowerCase() === name,
  );
  if (colIndex === -1) {
    throw new Error(`Column "${column}" not found in source range`);
  }
  return colIndex;
}

export function filterDataRows(
  rows: unknown[][],
  hasHeaders: boolean,
  filter?: RangeFilterSpec,
): { headerRow: unknown[] | null; filteredRows: unknown[][] } {
  const headerRow = hasHeaders && rows.length > 0 ? rows[0] : null;
  const dataRows = hasHeaders ? rows.slice(1) : rows;

  if (!filter) {
    return { headerRow, filteredRows: dataRows.map((row) => [...row]) };
  }

  if (!headerRow) {
    throw new Error('Filter requires hasHeaders: true so the column can be resolved');
  }

  const colIndex = resolveFilterColumnIndex(headerRow, filter.column);

  const filteredRows = dataRows
    .filter((row) => applyFilterOperator(row[colIndex], filter))
    .map((row) => [...row]);

  return { headerRow, filteredRows };
}

/** Row offsets within `rows` (including header offset) that match the filter. */
export function findMatchingRowOffsets(
  rows: unknown[][],
  hasHeaders: boolean,
  filter: RangeFilterSpec,
): number[] {
  if (rows.length === 0) return [];
  if (!hasHeaders) {
    throw new Error('findMatchingRowOffsets requires hasHeaders: true');
  }
  const headerRow = rows[0];
  const colIndex = resolveFilterColumnIndex(headerRow, filter.column);
  const matches: number[] = [];
  for (let i = 1; i < rows.length; i += 1) {
    if (applyFilterOperator(rows[i]?.[colIndex], filter)) {
      matches.push(i);
    }
  }
  return matches;
}

export function buildOutputRows(
  headerRow: unknown[] | null,
  filteredRows: unknown[][],
): unknown[][] {
  return headerRow ? [headerRow, ...filteredRows] : filteredRows;
}
