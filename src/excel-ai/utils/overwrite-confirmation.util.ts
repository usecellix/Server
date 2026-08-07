import { parseA1Range } from '../../agents/utils/range-merge.util';
import { stripSheetPrefix } from '../../agents/utils/range-address.util';

/**
 * Spec 21 — recognize legitimate overwrite confirmation so the frontend
 * overwrite guard does not block refinements of this conversation's own
 * recent writes (or explicit change/update language).
 */

export interface OverwriteTurnActionRecord {
  actionType: string;
  sheetName: string;
  /** Qualified or local A1 range this turn wrote into, e.g. "Purchase Register!L2:L51". */
  affectedRange?: string;
  /** Optional column name for SET_MATCHING_ROWS-style writes. */
  targetColumn?: string;
  turnIndex?: number;
}

/** Distinct from write-intent (01/11): confirms replacing already-populated content. */
const EXPLICIT_OVERWRITE_LANGUAGE =
  /\b(change|update|modify|correct|fix|replace|overwrite)\b[\s\S]{0,80}\b(to|with|as)\b/i;

export function hasExplicitOverwriteConfirmation(message: string): boolean {
  const text = String(message ?? '').trim();
  if (!text) return false;
  return EXPLICIT_OVERWRITE_LANGUAGE.test(text);
}

function splitQualifiedRange(range: string): { sheet: string | null; local: string } {
  const trimmed = String(range ?? '').trim().replace(/\$/g, '');
  if (!trimmed) return { sheet: null, local: '' };
  const bang = trimmed.lastIndexOf('!');
  if (bang < 0) return { sheet: null, local: trimmed };
  const sheet = trimmed
    .slice(0, bang)
    .replace(/^'+|'+$/g, '')
    .trim();
  return { sheet: sheet || null, local: trimmed.slice(bang + 1).trim() };
}

function sheetsCompatible(a: string | null, b: string | null): boolean {
  if (!a || !b) return true;
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

/**
 * True when two A1 ranges (optionally sheet-qualified) overlap in row/col space
 * on a compatible sheet.
 */
export function rangesOverlap(a: string, b: string): boolean {
  const left = splitQualifiedRange(a);
  const right = splitQualifiedRange(b);
  if (!left.local || !right.local) return false;
  if (!sheetsCompatible(left.sheet, right.sheet)) return false;

  const parsedA = parseA1Range(left.local);
  const parsedB = parseA1Range(right.local);
  if (!parsedA || !parsedB) return false;

  return !(
    parsedA.endRow < parsedB.startRow ||
    parsedB.endRow < parsedA.startRow ||
    parsedA.endCol < parsedB.startCol ||
    parsedB.endCol < parsedA.startCol
  );
}

/**
 * Target overlaps a range this conversation wrote in a recent turn.
 * Matches Spec 21 Signal 1 (refinement of own prior edit).
 *
 * Column-targeted priors (SET_MATCHING_ROWS) must match by targetColumn —
 * their stored table range is often the full used range and must not open
 * every column to overwrite confirmation.
 */
export function isRefinementOfOwnLastEdit(
  targetRange: string,
  lastNTurns: OverwriteTurnActionRecord[],
  opts?: { targetColumn?: string; sheetName?: string },
): boolean {
  const target = String(targetRange ?? '').trim();
  if (!target || !Array.isArray(lastNTurns) || lastNTurns.length === 0) {
    return false;
  }

  const targetCol = opts?.targetColumn?.trim().toLowerCase();
  const targetSheet = opts?.sheetName?.trim().toLowerCase();
  const targetParts = splitQualifiedRange(target);

  return lastNTurns.some((turn) => {
    const turnSheet = String(turn.sheetName ?? '').trim().toLowerCase();
    const sheetOk =
      !targetSheet ||
      !turnSheet ||
      targetSheet === turnSheet ||
      sheetsCompatible(targetParts.sheet, turnSheet || null);

    if (!sheetOk) return false;

    const priorCol = turn.targetColumn?.trim().toLowerCase();
    if (priorCol) {
      // Column-scoped prior write: require same named column.
      return Boolean(targetCol && targetCol === priorCol);
    }

    if (turn.affectedRange && rangesOverlap(turn.affectedRange, target)) {
      return true;
    }

    return false;
  });
}

function colIndexToLetter(index: number): string {
  let n = index + 1;
  let letter = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function qualifyRange(sheetName: string | undefined, local: string): string {
  const sheet = String(sheetName ?? '').trim();
  const range = stripSheetPrefix(local).trim();
  if (!range) return '';
  return sheet ? `${sheet}!${range}` : range;
}

function cellFromRowCol(row: number, col: number): string | null {
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) {
    return null;
  }
  return `${colIndexToLetter(col)}${row + 1}`;
}

/**
 * Resolve the write destination range for an action (for overlap / history).
 * Returns a sheet-qualified A1 string when possible.
 */
export function resolveActionWriteTarget(action: {
  type?: string;
  sheetName?: string;
  address?: string;
  row?: number;
  col?: number;
  range?: string;
  targetRange?: string;
  sourceRange?: string;
  destSheet?: string;
  destStartCell?: string;
  targetColumn?: string;
  operations?: Array<{ address?: string; row?: number; col?: number }>;
  headers?: unknown[];
  rows?: unknown[][];
  values?: unknown[];
  formula?: string;
  value?: unknown;
}): { sheetName: string; affectedRange: string; targetColumn?: string } | null {
  const type = String(action.type ?? '');

  if (type === 'SET_CELL' || type === 'SET_FORMULA') {
    const sheetName = String(action.sheetName ?? '').trim() || 'Sheet1';
    let local = '';
    if (typeof action.address === 'string' && action.address.trim()) {
      local = stripSheetPrefix(action.address);
    } else if (typeof action.row === 'number' && typeof action.col === 'number') {
      local = cellFromRowCol(action.row, action.col) ?? '';
    }
    if (!local) return null;
    return { sheetName, affectedRange: qualifyRange(sheetName, local) };
  }

  if (type === 'BATCH_SET') {
    const sheetName = String(action.sheetName ?? '').trim() || 'Sheet1';
    const cells: string[] = [];
    for (const op of action.operations ?? []) {
      if (typeof op.address === 'string' && op.address.trim()) {
        cells.push(stripSheetPrefix(op.address));
      } else if (typeof op.row === 'number' && typeof op.col === 'number') {
        const cell = cellFromRowCol(op.row, op.col);
        if (cell) cells.push(cell);
      }
    }
    if (cells.length === 0) return null;
    const parsed = cells
      .map((c) => parseA1Range(c))
      .filter((p): p is NonNullable<typeof p> => p != null);
    if (parsed.length === 0) return null;
    const startRow = Math.min(...parsed.map((p) => p.startRow));
    const startCol = Math.min(...parsed.map((p) => p.startCol));
    const endRow = Math.max(...parsed.map((p) => p.endRow));
    const endCol = Math.max(...parsed.map((p) => p.endCol));
    const local = `${colIndexToLetter(startCol)}${startRow + 1}:${colIndexToLetter(endCol)}${endRow + 1}`;
    return { sheetName, affectedRange: qualifyRange(sheetName, local) };
  }

  if (type === 'SET_MATCHING_ROWS' || type === 'FORMAT_MATCHING_ROWS') {
    const sheetName = String(action.sheetName ?? '').trim() || 'Sheet1';
    const local = stripSheetPrefix(String(action.range ?? action.sourceRange ?? '')).trim();
    if (!local) return null;
    const targetColumn =
      typeof action.targetColumn === 'string' && action.targetColumn.trim()
        ? action.targetColumn.trim()
        : undefined;
    return {
      sheetName,
      affectedRange: qualifyRange(sheetName, local),
      targetColumn,
    };
  }

  if (type === 'FILL_DOWN' || type === 'FILL_RIGHT' || type === 'AUTO_FILL') {
    const sheetName = String(action.sheetName ?? '').trim() || 'Sheet1';
    const local = stripSheetPrefix(
      String(action.targetRange ?? action.range ?? action.address ?? ''),
    ).trim();
    if (!local) return null;
    return { sheetName, affectedRange: qualifyRange(sheetName, local) };
  }

  if (type === 'WRITE_TABLE') {
    const sheetName = String(action.sheetName ?? '').trim() || 'Sheet1';
    const colCount = Math.max(
      Array.isArray(action.headers) ? action.headers.length : 0,
      ...(Array.isArray(action.rows)
        ? action.rows.map((row) => (Array.isArray(row) ? row.length : 0))
        : [0]),
      1,
    );
    const rowCount = (Array.isArray(action.headers) && action.headers.length ? 1 : 0) +
      (Array.isArray(action.rows) ? action.rows.length : 0);
    const endCol = colIndexToLetter(Math.max(colCount - 1, 0));
    const endRow = Math.max(rowCount, 1);
    return {
      sheetName,
      affectedRange: qualifyRange(sheetName, `A1:${endCol}${endRow}`),
    };
  }

  if (
    type === 'COPY_FILTERED_RANGE' ||
    type === 'MOVE_RANGE' ||
    type === 'AGGREGATE_TABLE'
  ) {
    const sheetName = String(action.destSheet ?? action.sheetName ?? '').trim() || 'Sheet1';
    const local = stripSheetPrefix(String(action.destStartCell ?? 'A1')).trim() || 'A1';
    return { sheetName, affectedRange: qualifyRange(sheetName, local) };
  }

  return null;
}

const OVERWRITE_ANNOTATABLE_TYPES = new Set([
  'SET_CELL',
  'SET_FORMULA',
  'BATCH_SET',
  'WRITE_TABLE',
  'FILL_DOWN',
  'FILL_RIGHT',
  'AUTO_FILL',
  'APPEND_ROW',
  'COPY_FILTERED_RANGE',
  'MOVE_RANGE',
  'AGGREGATE_TABLE',
  'SET_MATCHING_ROWS',
]);

type AnnotatableOverwriteAction = {
  type?: string;
  explicitOverwriteConfirmed?: boolean;
  sheetName?: string;
  targetColumn?: string;
};

/**
 * Set explicitOverwriteConfirmed when Spec 21 Signal 1 or 2 is present.
 * Does not clear an existing true flag.
 */
export function annotateExplicitOverwriteConfirmation<T extends AnnotatableOverwriteAction>(
  actions: T[],
  message: string,
  priorTurnActions: OverwriteTurnActionRecord[] = [],
): Array<T & { explicitOverwriteConfirmed?: boolean }> {
  if (!Array.isArray(actions) || actions.length === 0) return actions;

  const languageConfirmed = hasExplicitOverwriteConfirmation(message);

  return actions.map((action) => {
    if (action.explicitOverwriteConfirmed === true) return action;
    if (!OVERWRITE_ANNOTATABLE_TYPES.has(String(action.type ?? ''))) return action;

    if (languageConfirmed) {
      return { ...action, explicitOverwriteConfirmed: true };
    }

    const target = resolveActionWriteTarget(action as never);
    if (!target) return action;

    if (
      isRefinementOfOwnLastEdit(target.affectedRange, priorTurnActions, {
        targetColumn: target.targetColumn ?? action.targetColumn,
        sheetName: target.sheetName,
      })
    ) {
      return { ...action, explicitOverwriteConfirmed: true };
    }

    return action;
  });
}
