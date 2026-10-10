import { parseStatementDate, toIsoDate } from '../../domain-tools/ingestion/statement-values';

/**
 * A sheet read as a typed table, so a question about it can be answered by
 * code over every row (TASKS.md #380 to #383). The model never sees these
 * rows and never does the arithmetic.
 */

export type TableColumnType = 'number' | 'date' | 'text';

export interface TableColumn {
  index: number;
  name: string;
  type: TableColumnType;
  /** Decimal places to print for this column's numbers. */
  decimals: number;
}

export interface TableRow {
  /** 1-based row number on the sheet, for pointing the user at a row. */
  rowNumber: number;
  cells: unknown[];
}

export interface TableModel {
  sheetName: string;
  columns: TableColumn[];
  rows: TableRow[];
}

export interface ColumnHint {
  index: number;
  detectedType?: string;
  numberFormat?: string;
}

export interface BuildTableInput {
  sheetName: string;
  /** The sheet as read, header row included. */
  sheetData: unknown[][];
  /** 0-based index of the header row inside `sheetData`. */
  headerRowIndex?: number;
  headers?: string[];
  columnHints?: ColumnHint[];
}

export const isBlank = (value: unknown): boolean =>
  value === null || value === undefined || (typeof value === 'string' && value.trim() === '');

const NUMERIC_TEXT_RE = /^\(?[-+]?\s*(?:₹|rs\.?|inr|\$|£|€)?\s*[-+]?\d[\d,]*(?:\.\d+)?\s*%?\)?$/i;

/** A cell as a number, or null when it is not one. Accepts "1,25,000.50", "₹ 500", "(250)". */
export function toNumber(value: unknown): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value !== 'string') return null;
  const text = value.trim();
  if (!text || !NUMERIC_TEXT_RE.test(text)) return null;
  const negative = /^\(.*\)$/.test(text) || /^[^\d]*-/.test(text);
  const digits = text.replace(/[^\d.]/g, '');
  if (!digits) return null;
  const parsed = Number(digits);
  if (!Number.isFinite(parsed)) return null;
  return negative ? -parsed : parsed;
}

const EXCEL_EPOCH_OFFSET = 25_569; // days from 1899-12-30 to 1970-01-01
const MIN_DATE_SERIAL = 20_000; // 1954
const MAX_DATE_SERIAL = 80_000; // 2119

/** An Excel date serial as yyyy-mm-dd. The time of day, if any, is dropped. */
export function serialToIso(serial: number): string | null {
  if (!Number.isFinite(serial) || serial < 1 || serial > 2_958_465) return null;
  const date = new Date((Math.floor(serial) - EXCEL_EPOCH_OFFSET) * 86_400_000);
  return Number.isNaN(date.getTime()) ? null : date.toISOString().slice(0, 10);
}

/** A cell as a yyyy-mm-dd date, or null. Numbers are read as Excel serials. */
export function toIsoDateValue(value: unknown): string | null {
  if (typeof value === 'number') return serialToIso(value);
  if (typeof value !== 'string') return null;
  const parts = parseStatementDate(value);
  return parts ? toIsoDate(parts) : null;
}

/** True for a number format that displays a date ("dd-mmm-yyyy", "m/d/yy"), false for "#,##0.00" or "General". */
export function isDateNumberFormat(format: string | undefined): boolean {
  if (!format) return false;
  // Quoted text and colour/locale tags are not format codes.
  const code = format.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').toLowerCase();
  if (!code || code === 'general' || code === '@') return false;
  if (/[#0]/.test(code)) return false;
  return /[dy]/.test(code) || /m/.test(code);
}

function decimalsOf(value: number): number {
  if (Number.isInteger(value)) return 0;
  const text = String(value);
  if (text.includes('e')) return 6;
  return Math.min(6, text.split('.')[1]?.length ?? 0);
}

function formatDecimals(format: string | undefined): number | null {
  if (!format) return null;
  const match = /\.(0+)/.exec(format.replace(/"[^"]*"/g, ''));
  return match ? match[1].length : null;
}

const TYPE_SAMPLE_LIMIT = 400;

export function buildTable(input: BuildTableInput): TableModel {
  const headerRowIndex = Math.max(0, input.headerRowIndex ?? 0);
  const headerRow = input.sheetData[headerRowIndex] ?? [];
  const width = Math.max(
    headerRow.length,
    input.headers?.length ?? 0,
    ...input.sheetData.slice(headerRowIndex + 1, headerRowIndex + 51).map((row) => row?.length ?? 0),
  );

  const rows: TableRow[] = [];
  for (let r = headerRowIndex + 1; r < input.sheetData.length; r++) {
    const cells = input.sheetData[r] ?? [];
    if (!Array.isArray(cells) || cells.every(isBlank)) continue;
    rows.push({ rowNumber: r + 1, cells });
  }

  const usedNames = new Set<string>();
  const columns: TableColumn[] = [];
  for (let index = 0; index < width; index++) {
    const rawName = String(headerRow[index] ?? input.headers?.[index] ?? '').trim();
    let name = rawName || `Column ${index + 1}`;
    for (let n = 2; usedNames.has(name.toLowerCase()); n++) name = `${rawName || `Column ${index + 1}`} (${n})`;
    usedNames.add(name.toLowerCase());

    const hint = input.columnHints?.find((h) => h.index === index);
    const sample = rows.slice(0, TYPE_SAMPLE_LIMIT).map((row) => row.cells[index]).filter((v) => !isBlank(v));
    const numeric = sample.filter((v) => toNumber(v) !== null).length;
    const dateLike = sample.filter((v) => typeof v === 'string' && toIsoDateValue(v) !== null).length;
    const serialLike = sample.filter((v) => typeof v === 'number' && v >= MIN_DATE_SERIAL && v <= MAX_DATE_SERIAL).length;

    const hintedDate = hint?.detectedType === 'date' || isDateNumberFormat(hint?.numberFormat);
    let type: TableColumnType = 'text';
    if (sample.length > 0) {
      if (hintedDate && (serialLike + dateLike) / sample.length >= 0.6) type = 'date';
      else if (dateLike / sample.length >= 0.8) type = 'date';
      else if (numeric / sample.length >= 0.8) type = 'number';
    }

    let decimals = 0;
    if (type === 'number') {
      const fromFormat = formatDecimals(hint?.numberFormat);
      const seen = Math.max(0, ...sample.map((v) => decimalsOf(toNumber(v) ?? 0)));
      // Money keeps two places even when every amount in view is whole.
      decimals = fromFormat ?? (seen > 0 ? Math.max(2, Math.min(seen, 4)) : 0);
    }
    columns.push({ index, name, type, decimals });
  }

  return { sheetName: input.sheetName, columns, rows };
}

const norm = (text: string): string => text.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

/** Finds a column by the name the question used. Exact, then normalised, then a unique partial match. */
export function resolveColumn(table: TableModel, name: unknown): TableColumn | null {
  const wanted = String(name ?? '').trim();
  if (!wanted) return null;
  const exact = table.columns.find((c) => c.name === wanted);
  if (exact) return exact;
  const target = norm(wanted);
  if (!target) return null;
  const normalised = table.columns.filter((c) => norm(c.name) === target);
  if (normalised.length === 1) return normalised[0];
  // A shortened name ("Ref" for "Ref No") is accepted. A longer one is not:
  // "Date of birth" is not the Date column.
  const partial = table.columns.filter((c) => norm(c.name).includes(target));
  if (partial.length === 1) return partial[0];
  // A column letter ("E") is accepted when no header answers to it.
  if (/^[A-Za-z]{1,2}$/.test(wanted)) {
    let index = 0;
    for (const ch of wanted.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64);
    return table.columns[index - 1] ?? null;
  }
  return null;
}
