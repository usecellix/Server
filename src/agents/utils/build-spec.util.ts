import { SubTask } from '../types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 1 — the facts a user spelled out in
 * their prompt (which sheets, which columns, in what order), extracted once and
 * carried on the subtasks that write those headers.
 *
 * Without it the column list travels prompt -> Planner description -> Executor
 * as re-typed prose, and every hop can drift: a live run produced sheets whose
 * headers were placeholders ("Column1"...) or did not match what the user asked
 * for.
 */

export interface BuildSpecSheet {
  /** One name, or many for a repeated structure (12 month sheets share columns). */
  names: string[];
  /** The columns the user asked for, verbatim, in the order they gave them. */
  columns: string[];
}

export interface BuildSpec {
  sheets: BuildSpecSheet[];
}

export const normalizeHeaderText = (value: unknown): string =>
  typeof value === 'string'
    ? value
        .trim()
        .toLowerCase()
        .replace(/[_\-/]+/g, ' ')
        .replace(/[^\p{L}\p{N}\s]/gu, '')
        .replace(/\s+/g, ' ')
        .trim()
    : '';

const asStringArray = (value: unknown): string[] =>
  Array.isArray(value)
    ? value
        .filter((entry): entry is string => typeof entry === 'string')
        .map((entry) => entry.trim())
        .filter((entry) => entry.length > 0)
    : [];

/** Shape-checks whatever the extractor LLM returned; never throws. */
export function parseBuildSpec(raw: unknown): BuildSpec | null {
  if (!raw || typeof raw !== 'object') return null;
  const sheetsRaw = (raw as { sheets?: unknown }).sheets;
  if (!Array.isArray(sheetsRaw)) return null;

  const sheets: BuildSpecSheet[] = [];
  for (const entry of sheetsRaw) {
    if (!entry || typeof entry !== 'object') continue;
    const record = entry as { names?: unknown; name?: unknown; columns?: unknown };
    const names = asStringArray(record.names ?? (record.name ? [record.name] : []));
    const columns = asStringArray(record.columns);
    if (names.length > 0 && columns.length > 0) {
      sheets.push({ names, columns });
    }
  }
  return sheets.length > 0 ? { sheets } : null;
}

/**
 * Keeps only what the user actually wrote. An extractor that invents or
 * paraphrases a column would turn this safety net into the very drift it
 * exists to stop, so a column not present in the prompt drops the WHOLE sheet
 * entry (a partial column list would misdirect worse than none — the build then
 * simply falls back to today's behaviour for that sheet).
 */
export function groundBuildSpec(spec: BuildSpec, prompt: string): BuildSpec | null {
  const haystack = normalizeHeaderText(prompt);
  const sheets = spec.sheets.filter((sheet) => {
    if (sheet.columns.length < 2) return false;
    return sheet.columns.every((column) => {
      const needle = normalizeHeaderText(column);
      return needle.length > 0 && haystack.includes(needle);
    });
  });
  return sheets.length > 0 ? { sheets } : null;
}

/** Below this the plan is small enough that drift is not worth an extra LLM call. */
export const SPEC_MIN_SUBTASKS = 6;
export const SPEC_MIN_PROMPT_LENGTH = 200;

export function shouldExtractBuildSpec(prompt: string, subtasks: SubTask[]): boolean {
  return subtasks.length >= SPEC_MIN_SUBTASKS || prompt.length >= SPEC_MIN_PROMPT_LENGTH;
}

const HEADER_INTENT = /\bheaders?\b|\bcolumns?\b/i;

export const EXACT_HEADERS_MARKER = 'EXACT COLUMN HEADERS';

/**
 * Stamps `expectedHeaders` on every subtask that writes a spec'd sheet's header
 * row, and states them in the description the Executor actually reads.
 *
 * Only subtasks whose description mentions headers/columns are touched: a
 * formatting or formula pass over the same sheet must not be judged against a
 * header list it never writes.
 */
export function applyBuildSpecToSubtasks(subtasks: SubTask[], spec: BuildSpec): SubTask[] {
  return subtasks.map((subtask) => {
    if (subtask.expectedHeaders?.length) return subtask;
    if (!HEADER_INTENT.test(subtask.description)) return subtask;

    const target = subtask.targetSheet.trim().toLowerCase();
    const sheet = spec.sheets.find((entry) =>
      entry.names.some((name) => name.trim().toLowerCase() === target),
    );
    if (!sheet) return subtask;

    return {
      ...subtask,
      expectedHeaders: [...sheet.columns],
      description:
        `${subtask.description}\n\n${EXACT_HEADERS_MARKER} (verbatim, in this order — you may add ` +
        `computed columns after or between them, but never rename, drop or reorder these): ` +
        sheet.columns.join(' | '),
    };
  });
}

/**
 * Which of `expected` are missing or out of order in `written` (both in
 * left-to-right sheet order). Extra written columns are fine — the planner adds
 * computed columns (Nights, Total Amount) the user never listed.
 */
export function findHeaderMismatches(expected: string[], written: string[]): string[] {
  const writtenNorm = written.map(normalizeHeaderText);
  const missing: string[] = [];
  let cursor = 0;
  for (const column of expected) {
    const needle = normalizeHeaderText(column);
    const at = writtenNorm.indexOf(needle, cursor);
    if (at === -1) {
      missing.push(column);
    } else {
      cursor = at + 1;
    }
  }
  return missing;
}
