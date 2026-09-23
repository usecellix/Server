import { Action, SubTask } from '../types/agent.types';
import { buildHeaderTableActions, resolveHeaderRow } from './header-table-split.util';
import { findHeaderMismatches } from './build-spec.util';
import { isPlausibleSheetName } from './sheet-name.util';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 4 — diff what the run actually built
 * against what its own plan said it would, BEFORE telling the user it's done.
 *
 * Every failure mode this session chased ended the same way: the build
 * reported success while the workbook was missing sheets, missing header rows,
 * or carrying formulas that pointed at sheets nobody created. Each individual
 * cause got fixed; this is the backstop that catches the next one, whatever it
 * turns out to be, because it checks the OUTCOME rather than any particular
 * way of reaching it.
 *
 * Repairs are deterministic only. A missing sheet whose subtask carries
 * `expectedHeaders` can be rebuilt exactly (Phase 1.5's builder), so it is.
 * A dangling formula reference is REPORTED, never guessed at — inventing a
 * fix there is how you turn a visible gap into a silent wrong answer.
 */

export type ReconcileGapKind =
  | 'missing-sheet'
  | 'missing-header-row'
  | 'header-mismatch'
  | 'dangling-reference';

export interface ReconcileGap {
  kind: ReconcileGapKind;
  sheet: string;
  detail: string;
  /** True when `repairActions` below can close this gap with no LLM call. */
  repairable: boolean;
}

export interface ReconcileResult {
  gaps: ReconcileGap[];
  /** Deterministic actions that close every repairable gap, in apply order. */
  repairActions: Action[];
}

const key = (name: string): string => name.trim().toLowerCase();

const SHEET_REF = /(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!\$?[A-Za-z]{1,3}\$?\d*/g;
const ROW_ONE_CELL = /^\$?[A-Za-z]{1,3}\$?1$/;

function createdSheetName(action: Action): string | null {
  if (action.type !== 'ADD_SHEET' && action.type !== 'CREATE_SHEET') return null;
  const record = action as unknown as Record<string, unknown>;
  const name = String(record.name ?? record.sheetName ?? '').trim();
  return name || null;
}

/**
 * Sheets this action writes a row-1 (header) VALUE on.
 *
 * CREATE_TABLE is excluded deliberately: its range covers row 1 ("A1:E2") but
 * it writes nothing there — treating it as a header write is exactly what
 * would hide the Column1..N shape this check exists to catch.
 */
const NON_WRITING_TYPES = new Set([
  'CREATE_TABLE',
  'DELETE_TABLE',
  'FORMAT_RANGE',
  'SET_COLUMN_WIDTH',
  'DATA_VALIDATION',
  'SET_DATA_VALIDATION',
  'FREEZE_PANES',
  'CREATE_CHART',
]);

function writesRowOneOn(action: Action): string | null {
  if (NON_WRITING_TYPES.has(action.type)) return null;
  const record = action as unknown as Record<string, unknown>;
  const sheet = String(record.sheetName ?? '').trim();
  if (!sheet) return null;

  if (action.type === 'BATCH_SET' && Array.isArray(record.operations)) {
    const hit = (record.operations as Array<Record<string, unknown>>).some((op) =>
      ROW_ONE_CELL.test(String(op.address ?? '')),
    );
    return hit ? sheet : null;
  }
  const anchor = String(record.startCell ?? record.address ?? record.range ?? '');
  return /^\$?[A-Za-z]{1,3}\$?1(?::|$)/.test(anchor) ? sheet : null;
}

/** Row-1 header text this action wrote, left to right. */
function rowOneValuesOn(actions: Action[], sheet: string): string[] {
  const target = key(sheet);
  const cells: Array<{ col: number; text: string }> = [];
  const columnIndex = (letters: string): number =>
    letters.toUpperCase().split('').reduce((total, c) => total * 26 + (c.charCodeAt(0) - 64), 0);

  for (const action of actions) {
    const record = action as unknown as Record<string, unknown>;
    if (key(String(record.sheetName ?? '')) !== target) continue;
    if (action.type !== 'BATCH_SET' || !Array.isArray(record.operations)) continue;
    for (const op of record.operations as Array<Record<string, unknown>>) {
      const address = String(op.address ?? '');
      const match = /^\$?([A-Za-z]{1,3})\$?1$/.exec(address);
      if (match && typeof op.value === 'string' && !op.formula) {
        cells.push({ col: columnIndex(match[1]), text: op.value });
      }
    }
  }
  return cells.sort((a, b) => a.col - b.col).map((c) => c.text);
}

function referencedSheets(action: Action): string[] {
  const record = action as unknown as Record<string, unknown>;
  const texts: string[] = [];
  const collect = (v: unknown) => {
    if (typeof v === 'string' && v.startsWith('=')) texts.push(v);
  };
  collect(record.formula);
  if (Array.isArray(record.operations)) {
    for (const op of record.operations as Array<Record<string, unknown>>) {
      collect(op.formula);
      collect(op.value);
    }
  }
  const names = new Set<string>();
  for (const text of texts) {
    for (const m of text.matchAll(SHEET_REF)) {
      const name = (m[1]?.replace(/''/g, "'") ?? m[2] ?? '').trim();
      // TASKS.md #290 — a dynamically-built INDIRECT reference is not a sheet
      // name; reporting it as a dangling one would be a false alarm.
      if (name && isPlausibleSheetName(name)) names.add(name);
    }
  }
  return [...names];
}

export function reconcileRun(input: {
  subtasks: SubTask[];
  appliedActions: Action[];
  /** Sheets already in the workbook before this run started. */
  preExistingSheets?: string[];
}): ReconcileResult {
  // Defensive: this runs on the completion path of every stepwise build, so a
  // malformed or partially-hydrated run document must degrade to "nothing to
  // reconcile" rather than throw and take the finished build down with it.
  const subtasks = Array.isArray(input.subtasks) ? input.subtasks : [];
  const appliedActions = Array.isArray(input.appliedActions) ? input.appliedActions : [];
  const existing = new Set(
    (Array.isArray(input.preExistingSheets) ? input.preExistingSheets : [])
      .filter((name): name is string => typeof name === 'string')
      .map(key),
  );
  for (const action of appliedActions) {
    const created = createdSheetName(action);
    if (created) existing.add(key(created));
  }

  const gaps: ReconcileGap[] = [];
  const repairActions: Action[] = [];

  // 1. A sheet the plan targeted that nothing ever created.
  const seenTargets = new Set<string>();
  for (const subtask of subtasks) {
    const sheet = subtask.targetSheet?.trim();
    if (!sheet || seenTargets.has(key(sheet))) continue;
    seenTargets.add(key(sheet));
    if (existing.has(key(sheet))) continue;

    const headers = subtask.expectedHeaders?.length ? resolveHeaderRow(subtask) : [];
    const repairable = headers.length > 0;
    gaps.push({
      kind: 'missing-sheet',
      sheet,
      detail: repairable
        ? `Sheet "${sheet}" was planned but never created — rebuilding it with its ${headers.length} columns.`
        : `Sheet "${sheet}" was planned but never created, and carries no column list to rebuild it from.`,
      repairable,
    });
    if (repairable) {
      repairActions.push(...buildHeaderTableActions(subtask));
      existing.add(key(sheet));
    }
  }

  // 2. A sheet that exists but whose header row never got written — Excel
  //    names those columns "Column1".."ColumnN", the shape seen live.
  for (const subtask of subtasks) {
    const sheet = subtask.targetSheet?.trim();
    if (!sheet || !subtask.expectedHeaders?.length) continue;
    const wroteRowOne = appliedActions.some((a) => key(writesRowOneOn(a) ?? '') === key(sheet));
    if (wroteRowOne) continue;
    // Only meaningful if the sheet itself got created by this run.
    if (!appliedActions.some((a) => key(createdSheetName(a) ?? '') === key(sheet))) continue;

    const headers = resolveHeaderRow(subtask);
    gaps.push({
      kind: 'missing-header-row',
      sheet,
      detail: `Sheet "${sheet}" exists but its header row was never written.`,
      repairable: headers.length > 0,
    });
    if (headers.length > 0) {
      repairActions.push({
        type: 'BATCH_SET',
        sheetName: sheet,
        operations: headers.map((value, i) => ({
          address: `${String.fromCharCode(65 + (i % 26))}1`,
          value,
        })),
      } as Action);
    }
  }

  // 3. A header row that does not match what the user asked for.
  for (const subtask of subtasks) {
    const sheet = subtask.targetSheet?.trim();
    const expected = subtask.expectedHeaders;
    if (!sheet || !expected?.length) continue;
    const written = rowOneValuesOn(appliedActions, sheet);
    if (written.length === 0) continue; // covered by check 2
    const missing = findHeaderMismatches(expected, written);
    if (missing.length > 0) {
      gaps.push({
        kind: 'header-mismatch',
        sheet,
        detail: `Sheet "${sheet}" is missing or reordered column(s): ${missing.join(', ')}.`,
        repairable: false,
      });
    }
  }

  // 4. Formulas pointing at a sheet nothing ever created. Reported only —
  //    guessing a replacement reference would hide a real problem.
  const dangling = new Map<string, Set<string>>();
  for (const action of appliedActions) {
    const record = action as unknown as Record<string, unknown>;
    const onSheet = String(record.sheetName ?? '?');
    for (const ref of referencedSheets(action)) {
      if (existing.has(key(ref))) continue;
      const set = dangling.get(onSheet) ?? new Set<string>();
      set.add(ref);
      dangling.set(onSheet, set);
    }
  }
  for (const [sheet, refs] of dangling) {
    gaps.push({
      kind: 'dangling-reference',
      sheet,
      detail: `Formulas on "${sheet}" read from sheet(s) that do not exist: ${[...refs].join(', ')}.`,
      repairable: false,
    });
  }

  return { gaps, repairActions };
}

/** One honest sentence for the completion message. Empty when nothing is wrong. */
export function describeReconcileGaps(gaps: ReconcileGap[]): string {
  if (gaps.length === 0) return '';
  const repaired = gaps.filter((g) => g.repairable);
  const remaining = gaps.filter((g) => !g.repairable);

  const parts: string[] = [];
  if (repaired.length > 0) {
    parts.push(
      `Filled in ${repaired.length} missing piece${repaired.length === 1 ? '' : 's'} ` +
        `(${repaired.map((g) => g.sheet).join(', ')})`,
    );
  }
  if (remaining.length > 0) {
    parts.push(
      `${remaining.length} thing${remaining.length === 1 ? '' : 's'} still need attention: ` +
        remaining.map((g) => g.detail).join(' '),
    );
  }
  return parts.join('. ');
}
