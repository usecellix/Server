/**
 * Spec 23 — deterministic sheet overview for broad "tell me about this sheet" asks.
 * Aggregates are computed in code (buildAggregateTable / column sums), never by the LLM.
 */
import { buildAggregateTable } from '../../agents/utils/aggregate-table.util';
import { formatIndianCurrency, formatIndianNumber, parseIndianNumber } from './indian-format.util';
import type { SheetAnalysis } from '../services/sheet-analyzer.service';

const NUMERIC_HEADER_RE =
  /\b(amount|total|price|qty|quantity|tax|gst|cgst|sgst|igst|rate|value|spend|cost)\b/i;

const CATEGORICAL_HEADER_RE =
  /\b(status|type|category|payment|state|mode|kind|flag|class)\b/i;

const SUPPLIER_HEADER_RE = /\b(supplier|vendor|party|customer|client|name)\b/i;

const OVERVIEW_MESSAGE_RE =
  /\b(tell me about|describe|overview|summarize|summary of|what('?s| is) (in|on) (this|the)|what is in this|whats in this)\b/i;

export interface NumericSummary {
  column: string;
  sum: number;
  count: number;
}

export interface CategoryBreakdownRow {
  key: string;
  count: number;
  amount: number | null;
}

export interface CategoricalBreakdown {
  column: string;
  amountColumn?: string;
  rows: CategoryBreakdownRow[];
}

export interface TopRanking {
  categoryColumn: string;
  amountColumn: string;
  name: string;
  amount: number;
}

export interface SheetOverview {
  sheetName: string;
  rowCount: number;
  columnCount: number;
  columns: string[];
  dateRange: { column: string; from: string; to: string } | null;
  numericSummaries: NumericSummary[];
  categoricalBreakdowns: CategoricalBreakdown[];
  topRanking: TopRanking | null;
  dataQualityNotes: string[];
}

export function isSheetOverviewRequest(message: string): boolean {
  const lower = message.toLowerCase().trim();
  if (!lower) return false;
  if (OVERVIEW_MESSAGE_RE.test(lower)) return true;
  if (/\b(this|the)\s+(sheet|spreadsheet|workbook|data|table)\b/i.test(lower) && /\b(about|summar|overview|explain|describe)\b/i.test(lower)) {
    return true;
  }
  // Very short open-ended overview asks
  if (/^(about this sheet|sheet overview|summarize this|what'?s here\??)$/i.test(lower)) {
    return true;
  }
  return false;
}

/** Strip Spec 23 / Spec 15 internal vocabulary from LLM ask answers. */
export function sanitizeAskAnswer(text: string): string {
  if (!text) return text;
  let out = text;
  out = out.replace(/\bIntent\s*:\s*\w+/gi, '');
  out = out.replace(/\bdetectedType\s*:\s*[^\n,.;]+/gi, '');
  out = out.replace(/\bswitch to Action mode\b[^.!?\n]*/gi, '');
  // Flatten letter+type dumps like "A: number B: text" lines
  out = out.replace(
    /(?:^|\n)\s*(?:[A-Z]{1,3}\s*:\s*(?:number|text|date|string|boolean|empty)[^\n]*)+/gi,
    '\n',
  );
  out = out.replace(/[ \t]{2,}/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
  return out;
}

export function buildSheetOverview(
  sheetData: unknown[][],
  analysis: SheetAnalysis,
  sheetName = 'Sheet',
): SheetOverview {
  const headerRowIndex = analysis.headerRowIndex ?? 0;
  const tableRows = extractTableRows(sheetData, analysis);
  const dataRowCount = Math.max(0, tableRows.length - 1);
  const headers = analysis.headers.filter(Boolean);

  const numericCols = detectNumericColumns(tableRows, headers, dataRowCount);
  const amountCol =
    pickBestAmountColumn(headers, numericCols) ??
    (numericCols.length ? numericCols[numericCols.length - 1]!.name : undefined);

  const numericSummaries: NumericSummary[] = numericCols.map((col) => ({
    column: col.name,
    sum: col.sum,
    count: col.numericCount,
  }));

  const lowCard = detectLowCardinalityColumns(tableRows, headers, numericCols, dataRowCount);
  const categoricalBreakdowns: CategoricalBreakdown[] = [];

  for (const cat of lowCard.slice(0, 3)) {
    const breakdown = buildCategoryBreakdown(tableRows, cat.name, amountCol);
    if (breakdown.rows.length > 0) {
      categoricalBreakdowns.push(breakdown);
    }
  }

  const highCard = detectHighCardinalityCategories(tableRows, headers, numericCols, dataRowCount);
  let topRanking: TopRanking | null = null;
  if (amountCol && highCard.length > 0) {
    const catCol = highCard[0]!.name;
    try {
      const table = buildAggregateTable({
        rows: tableRows,
        hasHeaders: true,
        groupByColumn: catCol,
        aggregations: [{ column: amountCol, fn: 'sum', outputLabel: 'Total' }],
        sortBy: { column: 'Total', direction: 'desc' },
        topN: 1,
      });
      if (table.length >= 2) {
        const name = String(table[1]![0] ?? '').trim();
        const amount = toNum(table[1]![1]) ?? 0;
        if (name) {
          topRanking = {
            categoryColumn: catCol,
            amountColumn: amountCol,
            name,
            amount,
          };
        }
      }
    } catch {
      // column mismatch — skip ranking
    }
  }

  return {
    sheetName: sheetName || 'Sheet',
    rowCount: dataRowCount,
    columnCount: analysis.columnCount,
    columns: headers,
    dateRange: detectDateRange(tableRows, headers),
    numericSummaries,
    categoricalBreakdowns,
    topRanking,
    dataQualityNotes: detectDataQualityNotes(tableRows, headers, numericCols),
  };
}

export function formatSheetOverviewMarkdown(overview: SheetOverview): string {
  const blocks: string[] = [];
  const sheetLabel = overview.sheetName.replace(/\*\*/g, '');

  const datePart = overview.dateRange
    ? `, ${overview.dateRange.from} – ${overview.dateRange.to}`
    : '';
  blocks.push(`**${sheetLabel} overview**`);
  blocks.push(
    `- ${formatIndianNumber(overview.rowCount, 0)} records${datePart}`,
  );
  if (overview.columns.length) {
    blocks.push(`- Columns: ${overview.columns.join(', ')}`);
  }

  const qty = overview.numericSummaries.find((s) => /\b(qty|quantity)\b/i.test(s.column));
  const tax = overview.numericSummaries.find((s) =>
    /\b(tax amount|gst|cgst|sgst|igst)\b/i.test(s.column) && !/\b%/i.test(s.column),
  );
  const preTax = overview.numericSummaries.find((s) =>
    /\b(unit price|price|taxable|pre[- ]?tax|invoice amount|net)\b/i.test(s.column) &&
    !/\btotal\b/i.test(s.column),
  );
  const grand = overview.numericSummaries.find((s) =>
    /\btotal amount\b|\bgrand total\b|^total$/i.test(s.column),
  );

  const summaryBits: string[] = [];
  if (qty) summaryBits.push(`Total quantity: ${formatIndianNumber(qty.sum, 0)} units`);
  if (preTax) summaryBits.push(`Pre-tax: ${formatIndianCurrency(round2(preTax.sum))}`);
  if (tax) summaryBits.push(`GST: ${formatIndianCurrency(round2(tax.sum))}`);
  if (grand) summaryBits.push(`Total: ${formatIndianCurrency(round2(grand.sum))}`);

  if (summaryBits.length === 0 && overview.numericSummaries.length > 0) {
    for (const s of overview.numericSummaries.slice(0, 4)) {
      const isMoney = /\b(amount|total|price|tax|gst|cost|spend)\b/i.test(s.column);
      summaryBits.push(
        `${s.column}: ${isMoney ? formatIndianCurrency(round2(s.sum)) : formatIndianNumber(s.sum, 0)}`,
      );
    }
  }
  if (summaryBits.length) {
    blocks.push(`- ${summaryBits.join(' · ')}`);
  }

  for (const cat of overview.categoricalBreakdowns) {
    blocks.push('');
    blocks.push(`**${cat.column}**`);
    for (const row of cat.rows) {
      if (row.amount != null && cat.amountColumn) {
        blocks.push(
          `- ${row.key}: ${formatIndianNumber(row.count, 0)} records, ${formatIndianCurrency(round2(row.amount))}`,
        );
      } else {
        blocks.push(`- ${row.key}: ${formatIndianNumber(row.count, 0)} records`);
      }
    }
  }

  if (overview.topRanking) {
    blocks.push('');
    blocks.push(
      `**Largest ${overview.topRanking.categoryColumn.toLowerCase()}:** ${overview.topRanking.name}, ${formatIndianCurrency(round2(overview.topRanking.amount))}`,
    );
  }

  if (overview.dataQualityNotes.length) {
    blocks.push('');
    blocks.push('**Data quality notes**');
    for (const note of overview.dataQualityNotes) {
      blocks.push(`- ${note}`);
    }
  }

  blocks.push('');
  blocks.push('Want a totals row or a status breakdown added on the sheet?');

  return blocks.join('\n').trim();
}

// ── helpers ──────────────────────────────────────────────────────────

function extractTableRows(sheetData: unknown[][], analysis: SheetAnalysis): unknown[][] {
  const headerRowIndex = analysis.headerRowIndex ?? 0;
  const headers = analysis.headers;
  const dataBody = sheetData.slice(headerRowIndex + 1);
  // Prefer analyzer headers so group-by names match (even if row0 is a title).
  const headerRow = headers.map((h, i) => h || String(sheetData[headerRowIndex]?.[i] ?? `Column ${i + 1}`));
  return [headerRow, ...dataBody];
}

function toNum(value: unknown): number | null {
  return parseIndianNumber(value) ?? (typeof value === 'number' && Number.isFinite(value) ? value : null);
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function colSum(tableRows: unknown[][], colIndex: number): { sum: number; numericCount: number } {
  let sum = 0;
  let numericCount = 0;
  for (let r = 1; r < tableRows.length; r += 1) {
    const n = toNum(tableRows[r]?.[colIndex]);
    if (n === null) continue;
    sum += n;
    numericCount += 1;
  }
  return { sum, numericCount };
}

interface NumCol {
  name: string;
  index: number;
  sum: number;
  numericCount: number;
}

function detectNumericColumns(
  tableRows: unknown[][],
  headers: string[],
  dataRowCount: number,
): NumCol[] {
  const out: NumCol[] = [];
  for (let i = 0; i < headers.length; i += 1) {
    const name = headers[i]?.trim();
    if (!name) continue;
    const { sum, numericCount } = colSum(tableRows, i);
    const fraction = dataRowCount > 0 ? numericCount / dataRowCount : 0;
    const headerHit = NUMERIC_HEADER_RE.test(name);
    // Skip pure % rate columns for sum display when all values are 0–100 range — still allow if named tax amount
    if (/%\s*$/.test(name) && !/\bamount\b/i.test(name)) continue;
    if (headerHit || fraction >= 0.5) {
      if (numericCount === 0 && !headerHit) continue;
      out.push({ name, index: i, sum, numericCount: Math.max(numericCount, 0) });
    }
  }
  return out;
}

function pickBestAmountColumn(headers: string[], numericCols: NumCol[]): string | undefined {
  const preferred = numericCols.find((c) => /\btotal amount\b|\bgrand total\b/i.test(c.name));
  if (preferred) return preferred.name;
  const amount = numericCols.find((c) => /\bamount\b/i.test(c.name) && !/\btax\b/i.test(c.name));
  if (amount) return amount.name;
  const total = numericCols.find((c) => /\btotal\b/i.test(c.name));
  return total?.name;
}

function distinctCount(tableRows: unknown[][], colIndex: number): number {
  const set = new Set<string>();
  for (let r = 1; r < tableRows.length; r += 1) {
    const v = String(tableRows[r]?.[colIndex] ?? '').trim();
    if (v) set.add(v.toLowerCase());
  }
  return set.size;
}

function detectLowCardinalityColumns(
  tableRows: unknown[][],
  headers: string[],
  numericCols: NumCol[],
  dataRowCount: number,
): { name: string; index: number; distinct: number }[] {
  const numericIdx = new Set(numericCols.map((c) => c.index));
  const maxDistinct = Math.max(2, Math.min(12, Math.floor(dataRowCount / 3) || 2));
  const out: { name: string; index: number; distinct: number }[] = [];

  for (let i = 0; i < headers.length; i += 1) {
    if (numericIdx.has(i)) continue;
    const name = headers[i]?.trim();
    if (!name) continue;
    const d = distinctCount(tableRows, i);
    if (d < 2 || d > maxDistinct) continue;
    const preferred = CATEGORICAL_HEADER_RE.test(name) || d <= 6;
    if (!preferred && d > 6) continue;
    out.push({ name, index: i, distinct: d });
  }

  out.sort((a, b) => {
    const aPref = CATEGORICAL_HEADER_RE.test(a.name) ? 0 : 1;
    const bPref = CATEGORICAL_HEADER_RE.test(b.name) ? 0 : 1;
    if (aPref !== bPref) return aPref - bPref;
    return a.distinct - b.distinct;
  });
  return out;
}

function detectHighCardinalityCategories(
  tableRows: unknown[][],
  headers: string[],
  numericCols: NumCol[],
  dataRowCount: number,
): { name: string; index: number; distinct: number }[] {
  const numericIdx = new Set(numericCols.map((c) => c.index));
  const minDistinct = Math.max(3, Math.min(12, Math.floor(dataRowCount / 5) || 3));
  const out: { name: string; index: number; distinct: number }[] = [];

  for (let i = 0; i < headers.length; i += 1) {
    if (numericIdx.has(i)) continue;
    const name = headers[i]?.trim();
    if (!name) continue;
    const d = distinctCount(tableRows, i);
    if (d < minDistinct) continue;
    // Prefer supplier-like columns
    if (SUPPLIER_HEADER_RE.test(name) || d >= minDistinct) {
      out.push({ name, index: i, distinct: d });
    }
  }

  out.sort((a, b) => {
    const aPref = SUPPLIER_HEADER_RE.test(a.name) ? 0 : 1;
    const bPref = SUPPLIER_HEADER_RE.test(b.name) ? 0 : 1;
    if (aPref !== bPref) return aPref - bPref;
    return b.distinct - a.distinct;
  });
  return out;
}

function buildCategoryBreakdown(
  tableRows: unknown[][],
  categoryColumn: string,
  amountColumn?: string,
): CategoricalBreakdown {
  try {
    const table = buildAggregateTable({
      rows: tableRows,
      hasHeaders: true,
      groupByColumn: categoryColumn,
      aggregations: amountColumn
        ? [
            { column: amountColumn, fn: 'count', outputLabel: 'Count' },
            { column: amountColumn, fn: 'sum', outputLabel: 'Amount' },
          ]
        : [{ column: categoryColumn, fn: 'count', outputLabel: 'Count' }],
      sortBy: amountColumn
        ? { column: 'Amount', direction: 'desc' }
        : { column: 'Count', direction: 'desc' },
    });

    const rows: CategoryBreakdownRow[] = [];
    for (let r = 1; r < table.length; r += 1) {
      const key = String(table[r]![0] ?? '').trim();
      if (!key) continue;
      if (amountColumn) {
        rows.push({
          key,
          count: toNum(table[r]![1]) ?? 0,
          amount: toNum(table[r]![2]),
        });
      } else {
        rows.push({
          key,
          count: toNum(table[r]![1]) ?? 0,
          amount: null,
        });
      }
    }
    return { column: categoryColumn, amountColumn, rows };
  } catch {
    // Manual count fallback
    const header = tableRows[0] ?? [];
    const catIdx = header.findIndex(
      (h) => String(h).trim().toLowerCase() === categoryColumn.trim().toLowerCase(),
    );
    const amtIdx = amountColumn
      ? header.findIndex(
          (h) => String(h).trim().toLowerCase() === amountColumn.trim().toLowerCase(),
        )
      : -1;
    const map = new Map<string, { count: number; amount: number }>();
    for (let r = 1; r < tableRows.length; r += 1) {
      const key = String(tableRows[r]?.[catIdx] ?? '').trim();
      if (!key) continue;
      const entry = map.get(key) ?? { count: 0, amount: 0 };
      entry.count += 1;
      if (amtIdx >= 0) {
        const n = toNum(tableRows[r]?.[amtIdx]);
        if (n !== null) entry.amount += n;
      }
      map.set(key, entry);
    }
    const rows = [...map.entries()].map(([key, v]) => ({
      key,
      count: v.count,
      amount: amtIdx >= 0 ? v.amount : null,
    }));
    rows.sort((a, b) => (b.amount ?? b.count) - (a.amount ?? a.count));
    return { column: categoryColumn, amountColumn, rows };
  }
}

function parseDateLike(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
    const epoch = Date.UTC(1899, 11, 30);
    return new Date(epoch + value * 86400000);
  }
  if (typeof value === 'string' && value.trim()) {
    // Common dd-mm-yyyy / dd/mm/yyyy
    const m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/.exec(value.trim());
    if (m) {
      const d = Number(m[1]);
      const mo = Number(m[2]) - 1;
      let y = Number(m[3]);
      if (y < 100) y += 2000;
      const date = new Date(Date.UTC(y, mo, d));
      if (!Number.isNaN(date.getTime())) return date;
    }
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function formatDateShort(d: Date): string {
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const yyyy = d.getUTCFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

function detectDateRange(
  tableRows: unknown[][],
  headers: string[],
): { column: string; from: string; to: string } | null {
  for (let i = 0; i < headers.length; i += 1) {
    const name = headers[i]?.trim();
    if (!name) continue;
    let dateHits = 0;
    let min: Date | null = null;
    let max: Date | null = null;
    let textLikeDates = 0;
    for (let r = 1; r < tableRows.length; r += 1) {
      const raw = tableRows[r]?.[i];
      const d = parseDateLike(raw);
      if (!d) continue;
      dateHits += 1;
      if (typeof raw === 'string') textLikeDates += 1;
      if (!min || d < min) min = d;
      if (!max || d > max) max = d;
    }
    const dataRows = Math.max(1, tableRows.length - 1);
    if (dateHits >= Math.min(3, dataRows) && min && max) {
      return {
        column: name,
        from: formatDateShort(min),
        to: formatDateShort(max),
      };
    }
  }
  return null;
}

function detectDataQualityNotes(
  tableRows: unknown[][],
  headers: string[],
  numericCols: NumCol[],
): string[] {
  const notes: string[] = [];
  const dataRows = Math.max(0, tableRows.length - 1);

  for (let i = 0; i < headers.length; i += 1) {
    const name = headers[i]?.trim();
    if (!name) continue;

    // Date stored as text
    if (/\bdate\b/i.test(name)) {
      let textDates = 0;
      let dated = 0;
      for (let r = 1; r < tableRows.length; r += 1) {
        const raw = tableRows[r]?.[i];
        if (raw === null || raw === undefined || String(raw).trim() === '') continue;
        const d = parseDateLike(raw);
        if (d) {
          dated += 1;
          if (typeof raw === 'string') textDates += 1;
        }
      }
      if (dated > 0 && textDates >= dated * 0.8) {
        notes.push(
          `${name} values look like text dates — sorting and filtering by date may not work as expected.`,
        );
      }
    }
  }

  // Empty columns that look numeric from header
  for (const col of numericCols) {
    if (col.numericCount === 0 && dataRows > 0) {
      notes.push(
        `${col.name} looks like a numeric column but is currently empty — worth double-checking this is intentional.`,
      );
    }
  }

  // Fully empty columns by header name remarks
  for (let i = 0; i < headers.length; i += 1) {
    const name = headers[i]?.trim();
    if (!name || !/\bremark/i.test(name)) continue;
    let nonEmpty = 0;
    for (let r = 1; r < tableRows.length; r += 1) {
      const v = tableRows[r]?.[i];
      if (v !== null && v !== undefined && String(v).trim() !== '') nonEmpty += 1;
    }
    if (nonEmpty === 0 && dataRows > 0) {
      notes.push(
        `${name} is entirely empty — fill it in if you need annotations, or leave it blank intentionally.`,
      );
    }
  }

  return notes.slice(0, 5);
}
