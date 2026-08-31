import { SheetActionPayload } from '../types/sheet-actions.types';
import { WorkbookContext } from '../../types/cellix.types';

/**
 * Per-sheet answer to "is row 1 free to write?" — the only question the
 * header-row guard is actually asking.
 *
 * The guard exists to stop an LLM writing data over a user's real column
 * headers. That is a fact about one specific sheet, so it has to be resolved
 * per sheet. Grading a whole batch against the *active* sheet's emptiness
 * means a 13-sheet build is judged by whichever tab the user happened to have
 * open — and a batch that writes A1:J1 on twelve brand-new month sheets gets
 * merged into a single row. See TASKS.md #137.
 */
export type SheetHeaderStates = Map<string, boolean>;

const CREATE_TYPES = new Set<SheetActionPayload['type']>([
  'ADD_SHEET',
  'CREATE_SHEET',
  'COPY_SHEET',
]);

/** Normalized per-sheet key; '' means "the active sheet". */
export function sheetKeyOf(action: SheetActionPayload, activeSheet?: string): string {
  const explicit = String(action.sheetName ?? '').trim();
  if (explicit) return explicit.toLowerCase();
  return String(activeSheet ?? '').trim().toLowerCase();
}

/** Sheets this batch creates — empty by construction, so row 1 is free. */
export function sheetsCreatedInBatch(actions: SheetActionPayload[]): Set<string> {
  const created = new Set<string>();
  for (const action of actions) {
    if (!CREATE_TYPES.has(action.type)) continue;
    const name = String(
      action.type === 'COPY_SHEET'
        ? (action.newName ?? action.newSheetName ?? '')
        : (action.name ?? action.sheetName ?? ''),
    ).trim();
    if (name) created.add(name.toLowerCase());
  }
  return created;
}

/**
 * Resolve, for every sheet this batch touches, whether it already has a header
 * row. Truth comes from three sources, most authoritative first:
 *
 *  1. Created in this batch  → no header row (it does not exist yet).
 *  2. Present in the workbook context → its own headers/rowCount decide.
 *  3. Unknown → fall back to the caller's `activeSheetIsEmpty`, which is the
 *     pre-existing single-sheet behaviour.
 */
export function resolveSheetHeaderStates(
  actions: SheetActionPayload[],
  context: WorkbookContext | undefined,
  activeSheetIsEmpty: boolean,
): SheetHeaderStates {
  const states: SheetHeaderStates = new Map();
  const created = sheetsCreatedInBatch(actions);

  for (const key of created) states.set(key, false);

  for (const sheet of context?.sheets ?? []) {
    const key = sheet.sheetName.trim().toLowerCase();
    if (states.has(key)) continue;
    states.set(key, sheetHasHeaderRow(sheet.headers, sheet.rowCount));
  }

  const activeKey = String(context?.activeSheet ?? '').trim().toLowerCase();
  if (activeKey && !states.has(activeKey)) states.set(activeKey, !activeSheetIsEmpty);
  if (!states.has('')) {
    states.set('', activeKey ? (states.get(activeKey) ?? !activeSheetIsEmpty) : !activeSheetIsEmpty);
  }

  return states;
}

function sheetHasHeaderRow(headers: string[] | undefined, rowCount: number | undefined): boolean {
  const labelled = (headers ?? []).some((h) => String(h ?? '').trim() !== '');
  if (labelled) return true;
  // No labelled headers and no rows: nothing to protect.
  return (rowCount ?? 0) > 0 && labelled;
}

/**
 * Group actions by target sheet, preserving insertion order within each group.
 */
export function groupActionsBySheet(
  actions: SheetActionPayload[],
  activeSheet?: string,
): Map<string, SheetActionPayload[]> {
  const groups = new Map<string, SheetActionPayload[]>();
  for (const action of actions) {
    const key = sheetKeyOf(action, activeSheet);
    const group = groups.get(key);
    if (group) group.push(action);
    else groups.set(key, [action]);
  }
  return groups;
}
