import { WorkbookContext } from '../../types/cellix.types';

export interface SlicedSheetData {
  sheetName: string;
  headers: string[];
  columnIndices: number[];
  columnLetters: string[];
  rows: string[][];
  totalRows: number;
  headerRowIndex: number;
}

export interface SliceResult {
  sheets: SlicedSheetData[];
  /** columns that were mentioned but NOT found in any sheet */
  unresolved: string[];
}

/**
 * Extract only the columns likely needed to answer a data query from raw sheet rows.
 */
export function sliceFromRawData(
  message: string,
  rows: unknown[][],
  sheetName: string,
  knownHeaders?: string[],
): SliceResult {
  const normalizedRows = normalizeRows(rows);
  const keywords = extractColumnKeywords(message);
  const sliced = sliceSheet(normalizedRows, sheetName, keywords, knownHeaders);
  const unresolved = keywords.filter(
    (keyword) => !sliced.columnIndices.length || !matchesAnyHeader(keyword, sliced.headers),
  );

  return { sheets: [sliced], unresolved };
}

/**
 * Resolve active sheet name from workbook context and slice using supplied row data.
 */
export function sliceRelevantColumns(
  message: string,
  workbookContext: WorkbookContext | null | undefined,
  sheetData: unknown[][],
  activeSheetName?: string,
): SliceResult {
  if (!sheetData?.length) {
    const sheetName = resolveSheetName(workbookContext, activeSheetName);
    return { sheets: [emptySlice(sheetName)], unresolved: [] };
  }

  const sheetName = resolveSheetName(workbookContext, activeSheetName);
  const knownHeaders = resolveKnownHeaders(workbookContext, sheetName);
  return sliceFromRawData(message, sheetData, sheetName, knownHeaders);
}

function resolveSheetName(
  ctx: WorkbookContext | null | undefined,
  activeSheetName?: string,
): string {
  if (activeSheetName) {
    return activeSheetName;
  }
  if (ctx?.activeSheet) {
    return ctx.activeSheet;
  }
  if (ctx?.sheets?.length) {
    return ctx.sheets[0].sheetName;
  }
  return 'Sheet1';
}

function resolveKnownHeaders(
  ctx: WorkbookContext | null | undefined,
  sheetName: string,
): string[] | undefined {
  if (!ctx?.sheets?.length) {
    return undefined;
  }

  const snapshot =
    ctx.sheets.find((sheet) => sheet.sheetName.toLowerCase() === sheetName.toLowerCase()) ??
    ctx.sheets[0];
  const headers = snapshot.headers?.map((header) => String(header).trim()).filter(Boolean);
  return headers?.length ? headers : undefined;
}

function normalizeRows(rows: unknown[][]): string[][] {
  return rows.map((row) =>
    Array.isArray(row) ? row.map((cell) => cellToString(cell)) : [],
  );
}

function cellToString(cell: unknown): string {
  if (cell == null) {
    return '';
  }
  return String(cell).trim();
}

function extractColumnKeywords(message: string): string[] {
  const noise = new Set([
    'what',
    'is',
    'the',
    'total',
    'sum',
    'of',
    'in',
    'this',
    'sheet',
    'find',
    'show',
    'me',
    'all',
    'rows',
    'where',
    'how',
    'many',
    'count',
    'average',
    'mean',
    'max',
    'maximum',
    'min',
    'minimum',
    'get',
    'give',
    'calculate',
    'compute',
    'value',
    'values',
    'column',
    'columns',
    'a',
    'and',
    'or',
    'for',
    'with',
    'from',
    'to',
    'by',
    'than',
    'are',
    'have',
    'has',
    'that',
    'which',
    'do',
    'does',
    'please',
    'everything',
  ]);

  const taxTerms = [
    'cgst',
    'sgst',
    'igst',
    'tds',
    'tcs',
    'gst',
    'vat',
    'invoice',
    'voucher',
    'party',
    'amount',
    'taxable',
    'debit',
    'credit',
    'balance',
    'qty',
    'quantity',
    'rate',
    'date',
    'narration',
    'particulars',
  ];

  const lower = message.toLowerCase();
  const found: string[] = [];

  for (const term of taxTerms.sort((a, b) => b.length - a.length)) {
    if (messageIncludesTaxTerm(lower, term) && !found.some((existing) => existing.includes(term) && existing !== term)) {
      found.push(term);
    }
  }

  const words = lower
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter((word) => word.length > 2 && !noise.has(word));

  for (const word of words) {
    if (!found.includes(word)) {
      found.push(word);
    }
  }

  return [...new Set(found)];
}

function messageIncludesTaxTerm(lower: string, term: string): boolean {
  if (term.length <= 3) {
    return new RegExp(`\\b${term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i').test(lower);
  }
  return lower.includes(term);
}

function sliceSheet(
  rows: string[][],
  sheetName: string,
  keywords: string[],
  knownHeaders?: string[],
): SlicedSheetData {
  if (!rows.length) {
    return emptySlice(sheetName);
  }

  const headerRowIndex = detectHeaderRow(rows, knownHeaders);
  const headerRow = resolveHeaderRow(rows, headerRowIndex, knownHeaders);

  if (!keywords.length) {
    return buildSlice(
      sheetName,
      headerRow,
      rows,
      headerRowIndex,
      headerRow.map((_, index) => index),
    );
  }

  const matched = new Set<number>();

  for (const keyword of keywords) {
    for (let index = 0; index < headerRow.length; index++) {
      const header = (headerRow[index] ?? '').toLowerCase().trim();
      if (header.includes(keyword) || keyword.includes(header.replace(/\s+/g, ''))) {
        matched.add(index);
      }
    }
  }

  if (!matched.size) {
    return buildSlice(
      sheetName,
      headerRow,
      rows,
      headerRowIndex,
      headerRow.map((_, index) => index),
    );
  }

  addAnchorColumns(headerRow, matched);

  return buildSlice(
    sheetName,
    headerRow,
    rows,
    headerRowIndex,
    [...matched].sort((a, b) => a - b),
  );
}

function resolveHeaderRow(
  rows: string[][],
  headerRowIndex: number,
  knownHeaders?: string[],
): string[] {
  const row = rows[headerRowIndex] ?? [];
  if (knownHeaders?.length && knownHeaders.length >= row.length) {
    return knownHeaders.slice(0, Math.max(row.length, knownHeaders.length));
  }
  return row;
}

function addAnchorColumns(headers: string[], matched: Set<number>): void {
  const datePatterns = [/date/i, /dt/i];
  const keyPatterns = [/invoice/i, /voucher/i, /bill/i, /no\.?$/i, /sr\.?$/i, /id$/i];

  for (let index = 0; index < headers.length; index++) {
    if (matched.has(index)) {
      continue;
    }
    const header = headers[index] ?? '';
    if (datePatterns.some((pattern) => pattern.test(header)) || keyPatterns.some((pattern) => pattern.test(header))) {
      matched.add(index);
    }
  }
}

function buildSlice(
  sheetName: string,
  headerRow: string[],
  allRows: string[][],
  headerRowIndex: number,
  colIndices: number[],
): SlicedSheetData {
  const slicedHeaders = colIndices.map((index) => headerRow[index] ?? `Col${index + 1}`);
  const colLetters = colIndices.map((index) => columnIndexToLetter(index));
  const dataRows = allRows.slice(headerRowIndex + 1);
  const slicedRows = dataRows.map((row) => colIndices.map((index) => row[index] ?? ''));

  return {
    sheetName,
    headers: slicedHeaders,
    columnIndices: colIndices,
    columnLetters: colLetters,
    rows: slicedRows,
    totalRows: dataRows.length,
    headerRowIndex,
  };
}

function detectHeaderRow(rows: string[][], knownHeaders?: string[]): number {
  if (knownHeaders?.length) {
    const limit = Math.min(8, rows.length);
    for (let index = 0; index < limit; index++) {
      const row = rows[index] ?? [];
      const overlap = knownHeaders.filter(
        (header, headerIndex) =>
          header &&
          row[headerIndex] != null &&
          cellToString(row[headerIndex]).toLowerCase() === header.toLowerCase(),
      ).length;
      if (overlap >= Math.ceil(knownHeaders.length * 0.5)) {
        return index;
      }
    }
  }

  const limit = Math.min(8, rows.length);
  for (let index = 0; index < limit; index++) {
    const row = rows[index] ?? [];
    if (!row.length) {
      continue;
    }
    const textCells = row.filter((cell) => {
      const value = cell.trim();
      return value && Number.isNaN(Number(value.replace(/[,₹\s]/g, '')));
    }).length;
    if (textCells / row.length >= 0.6) {
      return index;
    }
  }
  return 0;
}

function matchesAnyHeader(keyword: string, headers: string[]): boolean {
  return headers.some(
    (header) =>
      header.toLowerCase().includes(keyword) || keyword.includes(header.toLowerCase()),
  );
}

function emptySlice(sheetName: string): SlicedSheetData {
  return {
    sheetName,
    headers: [],
    columnIndices: [],
    columnLetters: [],
    rows: [],
    totalRows: 0,
    headerRowIndex: 0,
  };
}

function columnIndexToLetter(index: number): string {
  let letter = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}
