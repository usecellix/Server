/**
 * P1.9 / A1.1 — "Clean the spaces and hidden characters in column A", deterministic lane.
 *
 * No LLM. Finds text cells with leading/trailing/double spaces, non-breaking spaces
 * (CHAR 160), tabs and other hidden characters, and proposes the cleaned text as an
 * ordinary BATCH_SET (so it still previews, needs Accept, and can be reverted).
 *
 * Never touched: numbers, dates, formula cells (reported as skipped), and any cell
 * whose cleaned text would be turned into a number when written (reported).
 */
import { BatchSetOperation, SheetActionPayload } from '../types/sheet-actions.types';
import { columnLetterOf, NBSP_RE, NON_PRINTING_RE } from './column-profiler.util';

export type CleanOp = 'trim' | 'collapseSpaces' | 'removeNbsp' | 'removeNonPrinting';
export const ALL_CLEAN_OPS: CleanOp[] = ['trim', 'collapseSpaces', 'removeNbsp', 'removeNonPrinting'];

const CLEAN_VERB = /\b(clean|cleanup|clean-up|trim|strip|remove|fix|get rid of|delete)\b/i;
const CLEAN_TARGET =
  /\b(spaces?|whitespace|white-space|hidden|non-?printing|invisible|nbsp|non-?breaking|extra spaces|control characters?|unwanted characters?)\b/i;
/** "delete blank rows", "remove duplicate rows" … are other jobs. */
const OTHER_JOB = /\b(rows?|duplicates?|blank cells?|columns?\s+(?:named|called)|sheets?\b.*\bnamed)\b/i;

export function detectCleanTextIntent(message: string): boolean {
  if (!message) return false;
  if (/\btrim\b/i.test(message) && !/\b(rows?|duplicates?)\b/i.test(message)) return true;
  return CLEAN_VERB.test(message) && CLEAN_TARGET.test(message) && !OTHER_JOB.test(message.replace(/\bin columns?\b/gi, ''));
}

/** Column letters named in the message: "column A", "columns A and C", "column A, B". */
export function cleanColumnsFromMessage(message: string): number[] | undefined {
  const m = /\bcolumns?\s+([A-Za-z]{1,2}(?:\s*(?:,|and|&)\s*[A-Za-z]{1,2})*)\b/i.exec(message);
  if (!m) return undefined;
  const letters = m[1].split(/\s*(?:,|and|&)\s*/i).filter((l) => /^[A-Za-z]{1,2}$/.test(l));
  if (letters.length === 0 || letters.some((l) => /^(of|in|is|to|or|for|the|has|are|an|as|on)$/i.test(l))) return undefined;
  return letters.map((l) => {
    let n = 0;
    for (const ch of l.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  });
}

/** Cleaned text for one cell value (strings only). */
export function cleanText(value: string, ops: CleanOp[] = ALL_CLEAN_OPS): string {
  let s = value;
  if (ops.includes('removeNbsp')) s = s.replace(/ /g, ' ');
  if (ops.includes('removeNonPrinting')) {
    s = s.replace(/\t/g, ' ').replace(NON_PRINTING_GLOBAL, '');
  }
  if (ops.includes('collapseSpaces')) s = s.replace(/ {2,}/g, ' ');
  if (ops.includes('trim')) s = s.trim();
  return s;
}

const NON_PRINTING_GLOBAL = new RegExp(NON_PRINTING_RE.source, 'g');

/** Spaces shown as · so invisible differences are visible in the preview. */
export function showSpaces(value: string): string {
  return value
    .replace(/ /g, '⍽')
    .replace(/\t/g, '→')
    .replace(NON_PRINTING_GLOBAL, '□')
    .replace(/ /g, '·');
}

/** A cleaned value Excel would silently turn into a number/date when written. */
function wouldBecomeNumber(cleaned: string): boolean {
  if (!cleaned) return false;
  if (/^[-+]?\d[\d,]*(\.\d+)?%?$/.test(cleaned)) return true;
  if (/^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(cleaned)) return true;
  return false;
}

export interface CleanTextCounts {
  columns: number;
  cellsChecked: number;
  toClean: number;
  alreadyClean: number;
  formulaSkipped: number;
  keptAsText: number;
}

export interface CleanTextResult {
  kind: 'actions' | 'answer';
  answer: string;
  explanation?: string;
  actions?: SheetActionPayload[];
  counts?: CleanTextCounts;
}

export interface CleanTextOptions {
  headerRowIndex?: number;
  /** Per column index → sheet row indexes (0-based, same indexing as `values`) that hold formulas. */
  formulaRows?: Array<number[] | undefined>;
  ops?: CleanOp[];
}

export function buildCleanTextPlan(
  values: unknown[][],
  sheetName: string,
  message: string,
  options: CleanTextOptions = {},
): CleanTextResult {
  const headerRowIndex = options.headerRowIndex ?? 0;
  const ops = options.ops ?? ALL_CLEAN_OPS;
  const width = values.reduce((w, r) => Math.max(w, Array.isArray(r) ? r.length : 0), 0);
  const requested = cleanColumnsFromMessage(message);
  const columns = requested ?? Array.from({ length: width }, (_, c) => c);

  const counts: CleanTextCounts = { columns: 0, cellsChecked: 0, toClean: 0, alreadyClean: 0, formulaSkipped: 0, keptAsText: 0 };
  const operations: BatchSetOperation[] = [];
  const samples: string[] = [];
  const perColumn: Array<{ letter: string; dirty: number }> = [];

  for (const c of columns) {
    if (c < 0 || c >= Math.max(width, 1)) continue;
    const formulaRows = new Set(options.formulaRows?.[c] ?? []);
    let dirty = 0;
    for (let r = headerRowIndex + 1; r < values.length; r += 1) {
      const v = values[r]?.[c];
      if (typeof v !== 'string' || v === '') continue; // numbers, dates, blanks: never touched
      counts.cellsChecked += 1;
      const cleaned = cleanText(v, ops);
      const isDirty = cleaned !== v;
      if (formulaRows.has(r)) {
        counts.formulaSkipped += 1;
        continue;
      }
      if (!isDirty) {
        counts.alreadyClean += 1;
        continue;
      }
      if (wouldBecomeNumber(cleaned)) {
        counts.keptAsText += 1;
        continue;
      }
      counts.toClean += 1;
      dirty += 1;
      const address = `${columnLetterOf(c)}${r + 1}`;
      operations.push({ address, value: cleaned });
      if (samples.length < 8) samples.push(`${address}: ${showSpaces(v)} → ${showSpaces(cleaned)}`);
    }
    if (dirty > 0) perColumn.push({ letter: columnLetterOf(c), dirty });
  }
  counts.columns = perColumn.length;

  const where = requested ? `column${requested.length > 1 ? 's' : ''} ${requested.map(columnLetterOf).join(', ')}` : 'this sheet';

  if (counts.toClean === 0) {
    const extras = [
      counts.formulaSkipped ? `${counts.formulaSkipped} formula cell${counts.formulaSkipped === 1 ? '' : 's'} skipped` : '',
      counts.keptAsText ? `${counts.keptAsText} cell${counts.keptAsText === 1 ? '' : 's'} left because cleaning would turn them into numbers` : '',
    ].filter(Boolean);
    return {
      kind: 'answer',
      answer: `Nothing to clean in ${where}: no extra spaces, non-breaking spaces or hidden characters found in ${counts.cellsChecked} text cell${counts.cellsChecked === 1 ? '' : 's'}.${extras.length ? ` (${extras.join('; ')}.)` : ''}`,
      counts,
    };
  }

  const lines = [
    `Cleaning text in ${where}: **${counts.toClean} cell${counts.toClean === 1 ? '' : 's'} to clean**` +
      (counts.alreadyClean ? `, ${counts.alreadyClean} already clean` : '') +
      (counts.formulaSkipped ? `, **${counts.formulaSkipped} formula cell${counts.formulaSkipped === 1 ? '' : 's'} skipped** (a formula's own text can't be edited in place)` : '') +
      '.',
    '',
    'What changes (spaces shown as ·, non-breaking space as ⍽, tab as →):',
    ...samples.map((s) => `- ${s}`),
  ];
  if (counts.toClean > samples.length) lines.push(`- …and ${counts.toClean - samples.length} more`);
  if (counts.keptAsText) {
    lines.push('', `${counts.keptAsText} cell${counts.keptAsText === 1 ? '' : 's'} left unchanged because cleaning would make Excel turn the text into a number or date.`);
  }
  lines.push('', '_Numbers, dates and formulas are never changed._');

  return {
    kind: 'actions',
    answer: lines.join('\n'),
    explanation: `Clean spaces and hidden characters in ${counts.toClean} cells.`,
    actions: [{ type: 'BATCH_SET', sheetName, operations }] as SheetActionPayload[],
    counts,
  };
}

export { NBSP_RE };
