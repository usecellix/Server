import { Action, SubTask } from '../types/agent.types';
import { findHeaderMismatches } from './build-spec.util';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md — Phase 1 pins the user's exact column list
 * onto a subtask (`expectedHeaders`), but that subtask still has to build the
 * WHOLE sheet in one Executor generation: create it, write the header row,
 * create the table, AND add formulas/dropdowns/widths/font. Two consecutive
 * live 12-month runs showed that single generation is too much work — even
 * running completely ALONE (no concurrency contention), the template subtask
 * hit "max iterations (10)" or timed out. Phase 2's concurrency/cloning work
 * was solving a different problem; this is the one actually blocking builds.
 *
 * Splitting a subtask's own PROSE reliably in two is fragile (different
 * prompts phrase the same request differently). What is NOT fragile is the
 * part we already have code-certain knowledge of: the header row and table
 * itself, entirely derivable from `expectedHeaders` and `targetSheet`. So the
 * create+headers+table piece is built by CODE with zero LLM calls (can't drift,
 * can't time out), and the ORIGINAL subtask keeps its full prose — now made
 * lighter because the biggest, most mechanical part of the work is already
 * done by the time it runs.
 */

const DEFAULT_ROW_HEIGHT_HEADER = 1;

function columnLetter(oneBasedIndex: number): string {
  let n = oneBasedIndex;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function columnIndexFromLetter(letters: string): number {
  return letters
    .toUpperCase()
    .split('')
    .reduce((total, char) => total * 26 + (char.charCodeAt(0) - 64), 0);
}

/** `<LETTER>1='<value>'` — one of the shapes the planner writes a header-row
 * instruction in (e.g. "A1='Unit No', B1='Guest', ..."). */
const HEADER_CELL_PATTERN = /\b([A-Z]{1,3})1\s*=\s*'([^']*)'/g;

/**
 * The OTHER shapes, all observed in real runs — a "headers" lead-in, an
 * optional `A1:M1` range hint, a separator, then a delimited list:
 *   "Write headers in row 1 (A1:M1): Unit No, Guest, Guest Name, ..."
 *   "headers in A1:M1 = [Unit No, Guest, Guest Name, ...]"
 *   "headers in row 1 — Unit No | Guest | Guest Name | ..."
 * TASKS.md #284 — #283 handled only the per-cell form, and the very next live
 * run used this one instead, so the parse fell back to the narrower
 * `expectedHeaders` and reopened the exact bug #283 was meant to close.
 */
/**
 * WITH a range hint. Tried first and matched explicitly, because the range's
 * OWN colon ("A1:M1") is otherwise mistaken for the list separator — which is
 * exactly what made the first cut of this parser return
 * `["M1): Unit No", "Guest", ...]`.
 */
const HEADER_LIST_WITH_RANGE =
  /headers?\b[^\n]{0,40}?\(?\b[A-Z]{1,3}1\s*:\s*([A-Z]{1,3})1\)?\s*[:=—-]\s*\[?([^.\n\]]+)\]?/i;

/** WITHOUT a range hint — "headers in row 1 — a | b | c". */
const HEADER_LIST_NO_RANGE = /headers?\b[^:=—\n]{0,40}?[:=—]\s*\[?([^.\n\]]+)\]?/i;

/** Splits "a, b, c" / "a | b | c" into trimmed, non-empty entries. */
function splitHeaderList(list: string): string[] {
  const delimiter = list.includes('|') ? '|' : ',';
  return list
    .split(delimiter)
    .map((entry) => entry.trim().replace(/^['"]|['"]$/g, '').replace(/[;,]$/, '').trim())
    .filter((entry) => entry.length > 0);
}

/**
 * The FULL header row the planner actually intended — the user's own columns
 * PLUS whatever computed ones it planned to interleave (Nights, Amount
 * Received, Balance Due...) — parsed straight out of the subtask's own
 * description. Returns null on anything that doesn't look like a complete
 * header row, so a stray "A1='Total'" KPI-cell mention is never mistaken for
 * one; `resolveHeaderRow` then falls back to the narrower, already-safe list.
 */
export function extractFullHeaderRowFromDescription(description: string): string[] | null {
  // Per-cell form first: it carries explicit positions, so it is the most
  // precise and can prove there are no gaps.
  const cells = new Map<number, string>();
  HEADER_CELL_PATTERN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = HEADER_CELL_PATTERN.exec(description))) {
    cells.set(columnIndexFromLetter(match[1]), match[2]);
  }
  if (cells.size >= 2) {
    const maxCol = Math.max(...cells.keys());
    const headers: string[] = [];
    for (let i = 1; i <= maxCol; i++) {
      const value = cells.get(i);
      if (value === undefined) return null; // a gap means this isn't a real header row
      headers.push(value);
    }
    return headers;
  }

  const withRange = HEADER_LIST_WITH_RANGE.exec(description);
  if (withRange) {
    const headers = splitHeaderList(withRange[2]);
    if (headers.length < 2) return null;
    // TASKS.md #289 — the range hint and the list can legitimately disagree.
    // A live planner wrote "In A1:J1 write headers: <13 columns> — that is 13
    // columns, so A1:M1", correcting itself mid-sentence. #284 required exact
    // equality, rejected that, and fell back to the narrower `expectedHeaders`
    // — which then contradicted the rest step's own 13-column instruction and
    // cost the run every month sheet's formulas.
    //
    // An explicit enumeration is stronger evidence than a range hint, so a
    // list LONGER than the hint is trusted. A list SHORTER than the hint is
    // still rejected: that is the truncated-split case the check was added
    // for, where the parse genuinely ran off. Either way `resolveHeaderRow`
    // independently verifies the result contains the user's own columns.
    if (headers.length < columnIndexFromLetter(withRange[1])) return null;
    return headers;
  }

  const noRange = HEADER_LIST_NO_RANGE.exec(description);
  if (!noRange) return null;
  const headers = splitHeaderList(noRange[1]);
  return headers.length >= 2 ? headers : null;
}

/**
 * Sentence-level removal of instructions the deterministic step has ALREADY
 * carried out — TASKS.md #284.
 *
 * #283 only PREPENDED "do not use ADD_SHEET/CREATE_TABLE/INSERT_COLUMN", and
 * the live run showed why that is not enough: the original description, still
 * embedded verbatim below that warning, went on to spell out the whole header
 * row and table anyway. Faced with one instruction saying "already built,
 * don't" and another saying "write these 13 headers", the Executor tried to
 * reconcile them and reached for INSERT_COLUMN. Removing the contradiction
 * outright is what actually settles it.
 *
 * Deliberately conservative: a chunk is dropped ONLY when it both starts with
 * a create-sheet/write-headers/create-table instruction AND carries no other
 * work (no formula, validation, width or font clause). Anything ambiguous is
 * left in — under-stripping costs a redundant instruction the prepend already
 * covers, while over-stripping would silently delete real work.
 */
const ALREADY_BUILT_SENTENCE =
  /^\s*(?:create\s+(?:the\s+|a\s+|an\s+)?sheet\b|write\s+(?:the\s+)?headers?\b|(?:set|add|put)\s+(?:the\s+)?headers?\b|create\s+(?:an\s+)?excel\s+table\b|create\s+table\b)/i;
const OTHER_WORK_KEYWORDS = /\bformula|validation|dropdown|width|font|freeze|format|chart|filter\b/i;

export function stripAlreadyBuiltInstructions(description: string): string {
  const chunks = description.split(/(?<=\.)\s+/);
  const kept = chunks.filter(
    (chunk) => !(ALREADY_BUILT_SENTENCE.test(chunk) && !OTHER_WORK_KEYWORDS.test(chunk)),
  );
  // Never strip everything — a description that was ONLY those instructions
  // still needs to say something, and an empty step reads as a silent drop.
  return kept.length > 0 ? kept.join(' ').trim() : description;
}

/**
 * The header row the deterministic step should actually build. Prefers the
 * planner's FULL intended layout (user columns + computed ones, in their
 * planned positions) over the bare `expectedHeaders` — TASKS.md #283: a live
 * run showed the cost of using `expectedHeaders` alone. The deterministic
 * step built only the user's 10 literal columns, but the "rest" step's own
 * (unmodified) description still described a 13-column layout with "Nights"
 * etc. inserted between them — the Executor, reconciling that mismatch,
 * tried to INSERT_COLUMN into a table that had no room for it ("target
 * column F already contains data"), and repeated failures left some sheets
 * with Excel's own placeholder headers mixed into the real ones.
 *
 * Falls back to `expectedHeaders` whenever the parsed full row doesn't
 * actually contain the user's own columns, in order — malformed or
 * unexpected phrasing must degrade to the narrower, already-safe behavior,
 * never risk building the WRONG header row from a bad parse.
 */
export function resolveHeaderRow(subtask: SubTask): string[] {
  // Resolved once at split time and carried on the step — see
  // `SubTask.resolvedHeaderRow`. TASKS.md #292.
  if (subtask.resolvedHeaderRow?.length) return subtask.resolvedHeaderRow;
  const expected = subtask.expectedHeaders ?? [];
  const full = extractFullHeaderRowFromDescription(subtask.description);
  if (full && findHeaderMismatches(expected, full).length === 0) {
    return full;
  }
  return expected;
}

/** `tbl` + the sheet name with everything but letters/digits stripped, e.g. "tblJanuary". */
export function deterministicTableName(sheetName: string): string {
  return `tbl${sheetName.replace(/[^A-Za-z0-9]/g, '')}`;
}

/**
 * ADD_SHEET + header row (BATCH_SET) + CREATE_TABLE (header + one blank data
 * row), built entirely from `resolveHeaderRow(subtask)`/`targetSheet` — no LLM
 * call, so it cannot time out, drift from the prompt, or hit an iteration cap.
 */
export function buildHeaderTableActions(subtask: SubTask): Action[] {
  const headers = resolveHeaderRow(subtask);
  const sheetName = subtask.targetSheet;
  const lastCol = columnLetter(headers.length);
  const tableName = deterministicTableName(sheetName);

  const addSheet: Action = { type: 'ADD_SHEET', name: sheetName, sheetName } as Action;
  // TASKS.md #289 — row 1 AND an empty row 2. The table below spans A1:<last>2,
  // but a row with no cells in it does not exist as far as the workbook context
  // is concerned: a live run reported `sheet 'February' has only 1 row
  // (A1:J1); row 2 does not exist, so F2/H2/L2 formulas ... cannot be written`
  // and every month's formula step was blocked. Seeding the data row is what
  // makes the table's own first row addressable by the step that follows.
  const writeHeaders: Action = {
    type: 'BATCH_SET',
    sheetName,
    operations: [
      ...headers.map((value, i) => ({ address: `${columnLetter(i + 1)}1`, value })),
      ...headers.map((_, i) => ({ address: `${columnLetter(i + 1)}2`, value: '' })),
    ],
  } as Action;
  const createTable: Action = {
    type: 'CREATE_TABLE',
    sheetName,
    range: `A1:${lastCol}${1 + DEFAULT_ROW_HEIGHT_HEADER}`,
    tableName,
    hasHeaders: true,
    showFilterButton: false,
  } as Action;

  return [addSheet, writeHeaders, createTable];
}

/**
 * Splits every Phase-1-pinned subtask (`expectedHeaders` present — i.e. it
 * writes a spec'd sheet's header row) into:
 *  - a new deterministic subtask that creates the sheet, header row and table,
 *  - the ORIGINAL subtask (same id, so every existing `dependsOn` reference to
 *    it — Main's consolidation subtasks, for instance — keeps working
 *    unmodified), now depending on the new step and told the sheet/headers/
 *    table already exist so it does not try to redo them.
 *
 * A subtask without `expectedHeaders` passes through untouched.
 *
 * The new id is derived from the SHEET NAME, not the original subtask id
 * (`hdr_January`, not `p2_s1_hdr`) — `template-replicate.util.ts`'s clone
 * detection normalizes sheet-name tokens (and only those) when comparing
 * subtasks across months, so the "rest" subtask's `dependsOn` (which now
 * includes this id) only stays clone-detectable if the id itself contains the
 * month name. An id built from the original subtask's own (month-unrelated)
 * index would silently opt every "rest" subtask back out of Phase 2 cloning.
 */
export function splitSpecPinnedSubtasks(subtasks: SubTask[]): SubTask[] {
  const usedIds = new Set(subtasks.map((s) => s.id));
  const result: SubTask[] = [];
  for (const subtask of subtasks) {
    if (!subtask.expectedHeaders?.length || subtask.isDeterministicHeaderStep) {
      result.push(subtask);
      continue;
    }

    const sheetSlug = subtask.targetSheet.replace(/[^A-Za-z0-9]/g, '') || subtask.id;
    let headerStepId = `hdr_${sheetSlug}`;
    let suffix = 2;
    while (usedIds.has(headerStepId)) {
      headerStepId = `hdr_${sheetSlug}_${suffix}`;
      suffix += 1;
    }
    usedIds.add(headerStepId);
    const tableName = deterministicTableName(subtask.targetSheet);
    // TASKS.md #283 — resolved ONCE and used for both the header step's own
    // actions (via `buildHeaderTableActions` re-deriving it the same way) and
    // the "rest" step's own range statement below, so the two can never
    // describe a different-sized table than the one that actually gets built.
    const fullHeaders = resolveHeaderRow(subtask);
    const headerStep: SubTask = {
      id: headerStepId,
      description:
        `Create sheet '${subtask.targetSheet}' with its header row and table '${tableName}' ` +
        `(built automatically — no further action needed here).`,
      targetSheet: subtask.targetSheet,
      dependsOn: subtask.dependsOn,
      estimatedActions: 3,
      expectedHeaders: subtask.expectedHeaders,
      // The row this step actually builds, decided here rather than re-derived
      // from the short description above — TASKS.md #292.
      resolvedHeaderRow: fullHeaders,
      isDeterministicHeaderStep: true,
    };

    const restStep: SubTask = {
      ...subtask,
      dependsOn: [...subtask.dependsOn, headerStepId],
      // Cleared, not carried forward: this subtask no longer writes the
      // header row, so `expectedHeaders` would be a stale claim about actions
      // it doesn't take — and would make a second pass over this same array
      // (e.g. from a caller that re-runs `applyBuildSpecToSubtasks`) split it
      // again, since the only other guard is `isDeterministicHeaderStep`.
      expectedHeaders: undefined,
      description:
        `Sheet '${subtask.targetSheet}', its header row (${fullHeaders.join(' | ')}) and table ` +
        `'${tableName}' over A1:${columnLetter(fullHeaders.length)}2 already exist (built ` +
        `automatically, in exactly this layout) — do NOT use ADD_SHEET, CREATE_TABLE or ` +
        `INSERT_COLUMN for this sheet, and do not rewrite the header row: every column below, ` +
        `including any computed ones, already has its header cell and column position. ` +
        `Add only what is still needed (formulas, validation, formatting):\n\n` +
        // TASKS.md #284 — the already-done create/header/table instructions are
        // REMOVED, not just warned against: leaving them in left the Executor
        // reconciling two contradictory instructions and reaching for
        // INSERT_COLUMN on a column that already had data.
        `${stripAlreadyBuiltInstructions(subtask.description)}`,
    };

    result.push(headerStep, restStep);
  }
  return result;
}
