import { Injectable } from '@nestjs/common';
import { parseIndianNumber } from '../utils/indian-format.util';

export interface SheetAnalysis {
  rowCount: number;
  columnCount: number;
  headers: string[];
  isEmpty: boolean;
  columnLetters: string[];
}

export interface ColumnStats {
  count: number;
  nonEmpty: number;
  sum: number;
  average: number;
  min: number | null;
  max: number | null;
  minRow: number | null;
  maxRow: number | null;
}

export interface DuplicateEntry {
  value: string;
  count: number;
  rows: number[];
}

@Injectable()
export class SheetAnalyzerService {
  analyze(sheetData: unknown[][]): SheetAnalysis {
    const rows = Array.isArray(sheetData) ? sheetData : [];
    const columnCount = rows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
    const headers = Array.isArray(rows[0])
      ? rows[0].map((cell, index) => this.formatHeader(cell, index))
      : [];

    return {
      rowCount: rows.length,
      columnCount,
      headers,
      isEmpty: rows.length === 0 || (rows.length === 1 && headers.every((h) => !h)),
      columnLetters: Array.from({ length: columnCount }, (_, i) => this.columnIndexToLetter(i)),
    };
  }

  getCellValue(sheetData: unknown[][], cellRef: string): unknown {
    const match = /^([A-Za-z]+)(\d+)$/.exec(cellRef.trim());
    if (!match) return undefined;

    const col = this.columnLetterToIndex(match[1].toUpperCase());
    const row = Number.parseInt(match[2], 10) - 1;
    if (row < 0 || col < 0) return undefined;

    const rowData = sheetData[row];
    if (!Array.isArray(rowData)) return undefined;
    return rowData[col];
  }

  sumColumn(sheetData: unknown[][], columnIndex: number, hasHeader = true): number {
    const startRow = hasHeader ? 1 : 0;
    let total = 0;

    for (let row = startRow; row < sheetData.length; row += 1) {
      const value = sheetData[row]?.[columnIndex];
      const numeric = this.toNumber(value);
      if (numeric !== null) total += numeric;
    }

    return total;
  }

  resolveColumnIndex(input: string, analysis: SheetAnalysis): number | null {
    const trimmed = input.trim();
    if (!trimmed) return null;

    const letterMatch = /^[A-Za-z]+$/.exec(trimmed);
    if (letterMatch) {
      return this.columnLetterToIndex(trimmed.toUpperCase());
    }

    const headerIndex = analysis.headers.findIndex(
      (header) => header.toLowerCase() === trimmed.toLowerCase(),
    );
    if (headerIndex >= 0) return headerIndex;

    const numeric = Number.parseInt(trimmed, 10);
    if (Number.isFinite(numeric) && numeric > 0 && numeric <= analysis.columnCount) {
      return numeric - 1;
    }

    return null;
  }

  resolveColumnIndexFromMessage(message: string, analysis: SheetAnalysis): number | null {
    const columnMatch = /\bcolumn\s+([A-Za-z0-9 ]+)\b/i.exec(message);
    if (columnMatch) {
      return this.resolveColumnIndex(columnMatch[1], analysis);
    }

    for (const header of analysis.headers) {
      if (header && message.toLowerCase().includes(header.toLowerCase())) {
        const idx = this.resolveColumnIndex(header, analysis);
        if (idx !== null) return idx;
      }
    }

    for (const letter of analysis.columnLetters) {
      const re = new RegExp(`\\b${letter}\\b`, 'i');
      if (re.test(message)) {
        return this.resolveColumnIndex(letter, analysis);
      }
    }

    return null;
  }

  columnStats(sheetData: unknown[][], columnIndex: number, hasHeader = true): ColumnStats {
    const startRow = hasHeader ? 1 : 0;
    let count = 0;
    let nonEmpty = 0;
    let sum = 0;
    let min: number | null = null;
    let max: number | null = null;
    let minRow: number | null = null;
    let maxRow: number | null = null;

    for (let row = startRow; row < sheetData.length; row += 1) {
      const raw = sheetData[row]?.[columnIndex];
      if (raw !== null && raw !== undefined && String(raw).trim() !== '') {
        nonEmpty += 1;
      }
      const num = this.toNumber(raw);
      if (num === null) continue;
      count += 1;
      sum += num;
      if (min === null || num < min) {
        min = num;
        minRow = row;
      }
      if (max === null || num > max) {
        max = num;
        maxRow = row;
      }
    }

    return {
      count,
      nonEmpty,
      sum,
      average: count > 0 ? sum / count : 0,
      min,
      max,
      minRow,
      maxRow,
    };
  }

  countBlank(sheetData: unknown[][], columnIndex: number, hasHeader = true): number {
    const startRow = hasHeader ? 1 : 0;
    let blank = 0;
    for (let row = startRow; row < sheetData.length; row += 1) {
      const value = sheetData[row]?.[columnIndex];
      if (value === null || value === undefined || String(value).trim() === '') {
        blank += 1;
      }
    }
    return blank;
  }

  countDataRows(sheetData: unknown[][], hasHeader = true): number {
    const startRow = hasHeader ? 1 : 0;
    let count = 0;
    for (let row = startRow; row < sheetData.length; row += 1) {
      const rowData = sheetData[row];
      if (Array.isArray(rowData) && rowData.some((c) => c !== null && c !== undefined && String(c).trim() !== '')) {
        count += 1;
      }
    }
    return count;
  }

  findDuplicates(sheetData: unknown[][], columnIndex: number, hasHeader = true): DuplicateEntry[] {
    const startRow = hasHeader ? 1 : 0;
    const map = new Map<string, number[]>();

    for (let row = startRow; row < sheetData.length; row += 1) {
      const value = sheetData[row]?.[columnIndex];
      if (value === null || value === undefined || String(value).trim() === '') continue;
      const key = String(value).trim();
      const rows = map.get(key) ?? [];
      rows.push(row);
      map.set(key, rows);
    }

    return [...map.entries()]
      .filter(([, rows]) => rows.length > 1)
      .map(([value, rows]) => ({ value, count: rows.length, rows }))
      .sort((a, b) => b.count - a.count);
  }

  findBlankRows(sheetData: unknown[][], columnIndex?: number, hasHeader = true): number[] {
    const startRow = hasHeader ? 1 : 0;
    const blankRows: number[] = [];

    for (let row = startRow; row < sheetData.length; row += 1) {
      const rowData = sheetData[row];
      if (!Array.isArray(rowData)) continue;

      if (columnIndex !== undefined) {
        const value = rowData[columnIndex];
        if (value === null || value === undefined || String(value).trim() === '') {
          blankRows.push(row);
        }
      } else if (rowData.every((c) => c === null || c === undefined || String(c).trim() === '')) {
        blankRows.push(row);
      }
    }

    return blankRows;
  }

  detectTextStoredNumbers(sheetData: unknown[][], columnIndex: number, hasHeader = true): boolean {
    const startRow = hasHeader ? 1 : 0;
    let textNumeric = 0;
    let realNumeric = 0;

    for (let row = startRow; row < Math.min(sheetData.length, startRow + 20); row += 1) {
      const raw = sheetData[row]?.[columnIndex];
      if (typeof raw === 'number') {
        realNumeric += 1;
      } else if (typeof raw === 'string' && parseIndianNumber(raw) !== null) {
        textNumeric += 1;
      }
    }

    return textNumeric > realNumeric && textNumeric > 0;
  }

  private formatHeader(cell: unknown, index: number): string {
    if (cell === null || cell === undefined || String(cell).trim() === '') {
      return `Column ${this.columnIndexToLetter(index)}`;
    }
    return String(cell);
  }

  private columnIndexToLetter(index: number): string {
    let n = index;
    let result = '';
    while (n >= 0) {
      result = String.fromCharCode((n % 26) + 65) + result;
      n = Math.floor(n / 26) - 1;
    }
    return result;
  }

  private columnLetterToIndex(letters: string): number {
    return letters.split('').reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
  }

  private toNumber(value: unknown): number | null {
    return parseIndianNumber(value);
  }
}
