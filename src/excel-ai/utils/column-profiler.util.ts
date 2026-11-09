/**
 * P1.5 / B1 — column profiler. Read-only: looks at a sheet's values (plus the
 * per-column number-format counts the add-in sends) and reports what is wrong
 * with each column. Deterministic, no LLM, never writes.
 *
 * Reused by: "Check the data quality of this sheet" (below), the cleaning
 * actions (P1.9 CLEAN_TEXT, colleague's C1.x/C2.x), formula diagnosis (P3.13).
 */

export interface FormatCount {
  format: string;
  count: number;
}

export type DateKind =
  | 'dd-mm-yyyy'
  | 'dd/mm/yyyy'
  | 'd.m.yy'
  | 'yyyy-mm-dd'
  | 'yyyymmdd'
  | 'dd-Mon-yyyy'
  | 'Mon dd, yyyy';

export interface ColumnProfile {
  index: number;
  letter: string;
  header: string;
  nonBlank: number;
  blanks: number;
  textNumbers: number;
  textDates: number;
  dateKinds: DateKind[];
  /** Leading/trailing/double ordinary spaces. */
  spacePadded: number;
  /** Non-breaking spaces (CHAR 160). */
  nbsp: number;
  /** Other non-printing characters (CHAR 0–31 except line breaks, zero-width, BOM). */
  nonPrinting: number;
  mixedCase: boolean;
  crFormatCells: number;
  drFormatCells: number;
  issues: string[];
}

export const NBSP_RE = / /;
export const NON_PRINTING_RE = /[\u0000-\u0008\u0009\u000B\u000C\u000E-\u001F\u007F​-‍⁠﻿]/;

export function columnLetterOf(index: number): string {
  let n = index;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

const isBlank = (v: unknown): boolean => v === null || v === undefined || String(v).trim() === '' ;

const TEXT_NUMBER_RE = /^[\s ]*[-+(]?\s*(?:₹|rs\.?|inr)?\s*\d[\d,]*(?:\.\d+)?\s*\)?\s*(?:cr|dr)?[\s ]*$/i;

/** True for a STRING cell that holds a number as text ("1,234", "₹ 7,500", "(2,500)"). */
export function isTextNumber(v: unknown): boolean {
  if (typeof v !== 'string') return false;
  const s = v.trim();
  if (!s || !TEXT_NUMBER_RE.test(s)) return false;
  // 8 digits that read as yyyymmdd are a date, not a number.
  if (dateKindOf(s) === 'yyyymmdd') return false;
  return true;
}

const MONTHS = 'jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec';

/** Which text date layout a STRING looks like, or null. */
export function dateKindOf(v: unknown): DateKind | null {
  if (typeof v !== 'string') return null;
  const s = v.trim();
  if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(s)) return 'dd-mm-yyyy';
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s)) return 'dd/mm/yyyy';
  if (/^\d{1,2}\.\d{1,2}\.\d{2,4}$/.test(s)) return 'd.m.yy';
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) return 'yyyy-mm-dd';
  if (/^(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$/.test(s)) return 'yyyymmdd';
  if (new RegExp(`^\\d{1,2}[-\\s](${MONTHS})[a-z]*[-\\s,]+\\d{2,4}$`, 'i').test(s)) return 'dd-Mon-yyyy';
  if (new RegExp(`^(${MONTHS})[a-z]*\\s+\\d{1,2},?\\s+\\d{4}$`, 'i').test(s)) return 'Mon dd, yyyy';
  return null;
}

type CasePattern = 'upper' | 'lower' | 'title' | 'mixed';

function casePatternOf(s: string): CasePattern | null {
  if (!/[A-Za-z]{2,}/.test(s)) return null;
  const letters = s.replace(/[^A-Za-z ]/g, '');
  if (letters === letters.toUpperCase()) return 'upper';
  if (letters === letters.toLowerCase()) return 'lower';
  const words = letters.split(/\s+/).filter(Boolean);
  if (words.every((w) => /^[A-Z][a-z]*$/.test(w))) return 'title';
  return 'mixed';
}

function hasDrCr(format: string): 'cr' | 'dr' | null {
  const m = /"?\b(cr|dr)\b"?/i.exec(format.replace(/\\/g, ''));
  return m ? (m[1].toLowerCase() as 'cr' | 'dr') : null;
}

export interface ProfileOptions {
  headerRowIndex?: number;
  /** Per column index → how many data cells use each number format (from the add-in). */
  formatCounts?: Array<FormatCount[] | undefined>;
  /** Restrict to these column indexes (e.g. "column D"). */
  onlyColumns?: number[];
}

export function profileColumns(values: unknown[][], opts: ProfileOptions = {}): ColumnProfile[] {
  const headerRowIndex = opts.headerRowIndex ?? 0;
  const header = values[headerRowIndex] ?? [];
  const width = values.reduce((w, r) => Math.max(w, Array.isArray(r) ? r.length : 0), 0);

  // Last row with any content, so trailing empty rows are not counted as blanks.
  let lastRow = headerRowIndex;
  for (let r = values.length - 1; r > headerRowIndex; r -= 1) {
    if (Array.isArray(values[r]) && values[r].some((v) => !isBlank(v))) {
      lastRow = r;
      break;
    }
  }

  const profiles: ColumnProfile[] = [];
  for (let c = 0; c < width; c += 1) {
    if (opts.onlyColumns && !opts.onlyColumns.includes(c)) continue;
    const p: ColumnProfile = {
      index: c,
      letter: columnLetterOf(c),
      header: String(header[c] ?? '').trim(),
      nonBlank: 0,
      blanks: 0,
      textNumbers: 0,
      textDates: 0,
      dateKinds: [],
      spacePadded: 0,
      nbsp: 0,
      nonPrinting: 0,
      mixedCase: false,
      crFormatCells: 0,
      drFormatCells: 0,
      issues: [],
    };
    const kinds = new Set<DateKind>();
    const casePatterns = new Map<CasePattern, number>();
    let textCells = 0;

    for (let r = headerRowIndex + 1; r <= lastRow; r += 1) {
      const v = values[r]?.[c];
      if (isBlank(v)) {
        p.blanks += 1;
        continue;
      }
      p.nonBlank += 1;
      if (typeof v !== 'string') continue;
      textCells += 1;
      const hasNbsp = NBSP_RE.test(v);
      const hasOther = NON_PRINTING_RE.test(v);
      if (hasNbsp) p.nbsp += 1;
      if (hasOther) p.nonPrinting += 1;
      const plain = v.replace(/ /g, ' ');
      if (plain !== plain.trim() || /\s{2,}/.test(plain)) p.spacePadded += 1;
      if (isTextNumber(v)) p.textNumbers += 1;
      const kind = dateKindOf(v);
      if (kind) {
        p.textDates += 1;
        kinds.add(kind);
      }
      const pattern = casePatternOf(v.trim());
      if (pattern) casePatterns.set(pattern, (casePatterns.get(pattern) ?? 0) + 1);
    }

    p.dateKinds = [...kinds];

    // Mixed case: ≥2 case patterns and the second-most-common is ≥25% of text values.
    if (casePatterns.size >= 2 && textCells >= 2) {
      const counts = [...casePatterns.values()].sort((a, b) => b - a);
      p.mixedCase = counts[1] / textCells >= 0.25;
    }

    for (const fc of opts.formatCounts?.[c] ?? []) {
      const kind = hasDrCr(fc.format);
      if (kind === 'cr') p.crFormatCells += fc.count;
      if (kind === 'dr') p.drFormatCells += fc.count;
    }

    // ---- issue sentences (kept short; counts are exact) ----
    // "Text-stored numbers" only when most of the filled cells are numbers-as-text,
    // so an ID column with one stray "0233" is not mislabelled.
    if (p.textNumbers > 0 && p.textNumbers / Math.max(p.nonBlank, 1) >= 0.5) {
      p.issues.push(`${p.textNumbers} text-stored number${p.textNumbers === 1 ? '' : 's'} (look like numbers but are text, so SUM/lookups skip or miss them)`);
    }
    if (p.crFormatCells + p.drFormatCells > 0) {
      p.issues.push(
        `Dr/Cr is in the cell number format, not the value (${p.crFormatCells} Cr, ${p.drFormatCells} Dr) — the cell holds a positive number either way`,
      );
    }
    if (p.dateKinds.length >= 2) {
      p.issues.push(`${p.dateKinds.length} different text date formats (${p.dateKinds.join(', ')})`);
    } else if (p.dateKinds.length === 1 && p.textDates >= Math.max(p.nonBlank * 0.5, 1)) {
      p.issues.push(`${p.textDates} date${p.textDates === 1 ? '' : 's'} stored as text (${p.dateKinds[0]})`);
    }
    if (p.spacePadded > 0) p.issues.push(`${p.spacePadded} cell${p.spacePadded === 1 ? '' : 's'} with leading, trailing or double spaces`);
    if (p.nbsp > 0) p.issues.push(`${p.nbsp} cell${p.nbsp === 1 ? '' : 's'} with non-breaking spaces (CHAR 160)`);
    if (p.nonPrinting > 0) p.issues.push(`${p.nonPrinting} cell${p.nonPrinting === 1 ? '' : 's'} with hidden characters (tab / control)`);
    if (p.mixedCase) p.issues.push('mixed upper / lower / title case');
    if (p.blanks > 0 && p.nonBlank > 0) p.issues.push(`${p.blanks} blank cell${p.blanks === 1 ? '' : 's'}`);

    profiles.push(p);
  }
  return profiles;
}

const HEADER_COL_RE = /\bcolumns?\s+([A-Za-z]{1,2})\b/i;

/** "Check the data quality…", "What's wrong with column D?" — read-only report intents. */
export function detectProfileIntent(message: string): boolean {
  if (!message) return false;
  if (/formula|#n\/a|#ref|#value|error/i.test(message)) return false;
  return (
    /\bdata[-\s]?quality\b/i.test(message) ||
    /\b(what'?s|what is|whats)\s+wrong\s+with\b[\s\S]*\b(columns?|sheet|data|table)\b/i.test(message) ||
    /\b(profile|health[-\s]?check)\b[\s\S]*\b(columns?|sheet|data|table)\b/i.test(message) ||
    /\bany\s+(data\s+)?(issues|problems)\b[\s\S]*\b(columns?|sheet|data)\b/i.test(message)
  );
}

export function profileOnlyColumnsFromMessage(message: string): number[] | undefined {
  const m = HEADER_COL_RE.exec(message);
  if (!m || /^(of|in|is|to|and|or|for|the|has|are)$/i.test(m[1])) return undefined;
  let n = 0;
  for (const ch of m[1].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return [n - 1];
}

export function buildProfileReport(
  sheetName: string,
  profiles: ColumnProfile[],
): { answer: string; flaggedColumns: number } {
  const flagged = profiles.filter((p) => p.issues.some((i) => !/blank cell/.test(i)) || p.blanks > 0 && p.nonBlank > 0);
  const lines: string[] = [];
  if (flagged.length === 0) {
    lines.push(`**${sheetName}** looks clean: no text-stored numbers, hidden characters, mixed date formats or Dr/Cr number formats found in ${profiles.length} column${profiles.length === 1 ? '' : 's'}.`);
  } else {
    lines.push(`**Data-quality check of ${sheetName}** — ${flagged.length} of ${profiles.length} column${profiles.length === 1 ? '' : 's'} need attention. Nothing was changed.`);
    lines.push('');
    for (const p of flagged) {
      const name = p.header ? `${p.letter} (${p.header})` : p.letter;
      lines.push(`**Column ${name}**`);
      for (const issue of p.issues) lines.push(`- ${issue}`);
    }
  }
  const hints: string[] = [];
  if (flagged.some((p) => p.spacePadded || p.nbsp || p.nonPrinting)) hints.push('*Clean the spaces and hidden characters in column X*');
  if (hints.length) lines.push('', `Next step you can ask for: ${hints.join(' · ')}`);
  return { answer: lines.join('\n'), flaggedColumns: flagged.length };
}
