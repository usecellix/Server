import {
  RangeFilterOperator,
  RangeFilterSpec,
} from '../../excel-ai/types/sheet-actions.types';

import { isTextNumber } from '../../excel-ai/utils/column-profiler.util';
import { validateGstin } from '../../excel-ai/utils/india-validators.util';

export type { RangeFilterSpec };

/** Mirrors client/src/shared/rangeFilter.ts (P1.7) so the shadow workbook predicts the same rows the add-in will touch. */
function parseDateMs(cell: unknown): number | null {
  if (typeof cell === 'number' && Number.isFinite(cell)) {
    return cell > 20_000 && cell < 80_000 ? Math.round((cell - 25_569) * 86_400_000) : null;
  }
  if (typeof cell !== 'string') return null;
  const t = cell.trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(t);
  if (m) {
    const year = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    const month = +m[2];
    const day = +m[1];
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    return Date.UTC(year, month - 1, day);
  }
  return null;
}

function findDuplicateOffsets(dataRows: unknown[][], keyIndexes: number[], keep: 'first' | 'last' = 'first'): number[] {
  const keyOf = (row: unknown[]): string | null => {
    const parts = keyIndexes.map((i) => String(row?.[i] ?? '').trim().toLowerCase());
    return parts.every((part) => part === '') ? null : parts.join('\u0001');
  };
  const order = keep === 'first' ? dataRows.map((_, i) => i) : dataRows.map((_, i) => dataRows.length - 1 - i);
  const seen = new Set<string>();
  const dupes: number[] = [];
  for (const i of order) {
    const key = keyOf(dataRows[i]);
    if (key === null) continue;
    if (seen.has(key)) dupes.push(i);
    else seen.add(key);
  }
  return dupes.sort((a, b) => a - b);
}

export function applyFilterOperator(
  cellValue: unknown,
  filter: RangeFilterSpec,
): boolean {
  const { operator, value } = filter;

  if (operator === 'isBlank' || operator === 'isNotBlank') {
    const trim = String(value ?? '') === 'trimSpaces';
    const blank = cellValue === null || cellValue === undefined || (trim ? String(cellValue).trim() === '' : String(cellValue) === '');
    return operator === 'isBlank' ? blank : !blank;
  }
  if (operator === 'isTextNumber') return isTextNumber(cellValue);
  if (operator === 'isInvalidGstin') {
    return String(cellValue ?? '').trim() !== '' && validateGstin(cellValue).status !== 'valid';
  }
  if (operator === 'olderThan') {
    const ms = parseDateMs(cellValue);
    const days = Number(value);
    if (ms === null || !Number.isFinite(days)) return false;
    const ref = filter.asOn ? parseDateMs(filter.asOn) : null;
    const now = new Date();
    const refMs = ref ?? Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    return (refMs - ms) / 86_400_000 > days;
  }
  if (operator === 'isDuplicate') return false; // row-level - handled in findMatchingRowOffsets / filterDataRows

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

  if (filter.operator === 'isDuplicate') {
    const keys = (filter.keyColumns?.length ? filter.keyColumns : [filter.column]).map((c) => resolveFilterColumnIndex(headerRow, c));
    const dupes = new Set(findDuplicateOffsets(dataRows, keys, filter.keep));
    return { headerRow, filteredRows: dataRows.filter((_, i) => dupes.has(i)).map((row) => [...row]) };
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
  if (filter.operator === 'isDuplicate') {
    const keys = (filter.keyColumns?.length ? filter.keyColumns : [filter.column]).map((c) => resolveFilterColumnIndex(headerRow, c));
    return findDuplicateOffsets(rows.slice(1), keys, filter.keep).map((i) => i + 1);
  }
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
