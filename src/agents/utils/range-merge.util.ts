import { SheetContext } from '../types/agent.types';

export interface ParsedRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export function columnLetterToIndex(letters: string): number {
  let index = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i += 1) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function parseA1Cell(cell: string): { row: number; col: number } | null {
  const match = /^([A-Za-z]+)(\d+)$/.exec(cell.trim());
  if (!match) return null;
  return {
    col: columnLetterToIndex(match[1]),
    row: Number.parseInt(match[2], 10) - 1,
  };
}

export function parseA1Range(range: string): ParsedRange | null {
  const trimmed = range.trim().replace(/\$/g, '');
  const parts = trimmed.split(':');
  const start = parseA1Cell(parts[0] ?? '');
  if (!start) return null;

  const end = parts.length > 1 ? parseA1Cell(parts[1] ?? '') : start;
  if (!end) return null;

  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

function ensureRow(matrix: unknown[][], rowIndex: number, columnCount: number): void {
  while (matrix.length <= rowIndex) {
    matrix.push(Array.from({ length: columnCount }, () => ''));
  }
  const row = matrix[rowIndex];
  if (!Array.isArray(row)) {
    matrix[rowIndex] = Array.from({ length: columnCount }, () => '');
    return;
  }
  while (row.length < columnCount) {
    row.push('');
  }
}

export function mergeRangeIntoSheet(
  sheet: SheetContext,
  range: string,
  values: unknown[][],
): SheetContext {
  const parsed = parseA1Range(range);
  if (!parsed) return sheet;

  const columnCount = Math.max(
    sheet.columnCount,
    parsed.endCol + 1,
    values.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0),
  );

  const newValues = sheet.values.map((row) =>
    Array.isArray(row) ? [...row] : Array.from({ length: columnCount }, () => ''),
  );
  const newFormulas = sheet.formulas.map((row) =>
    Array.isArray(row) ? [...row] : Array.from({ length: columnCount }, () => ''),
  );

  for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
    const targetRow = parsed.startRow + rowOffset;
    const sourceRow = values[rowOffset];
    if (!Array.isArray(sourceRow)) continue;

    ensureRow(newValues, targetRow, columnCount);
    ensureRow(newFormulas, targetRow, columnCount);

    for (let colOffset = 0; colOffset < sourceRow.length; colOffset += 1) {
      const targetCol = parsed.startCol + colOffset;
      const cell = sourceRow[colOffset];
      if (typeof cell === 'string' && cell.startsWith('=')) {
        newFormulas[targetRow][targetCol] = cell;
        newValues[targetRow][targetCol] = cell;
      } else {
        newValues[targetRow][targetCol] = cell ?? '';
        if (!newFormulas[targetRow][targetCol]) {
          newFormulas[targetRow][targetCol] = '';
        }
      }
    }
  }

  const rowCount = Math.max(sheet.rowCount, parsed.endRow + 1, newValues.length);

  return {
    ...sheet,
    rowCount,
    columnCount,
    values: newValues,
    formulas: newFormulas,
  };
}
