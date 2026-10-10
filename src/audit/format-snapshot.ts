import { Action } from '../agents/types/agent.types';
import { RevertNoOpError } from './errors/revert-noop.error';

/**
 * Revert for formatting (TASKS.md #400).
 *
 * A formatting action (FORMAT_RANGE, AUTOFIT_COLUMNS) changes nothing the shadow
 * workbook simulates, so a change set has no "before" for it. The add-in reads
 * the previous state straight off Excel just before it applies each of those
 * actions and reports it with the apply call; it is stored on the change set,
 * and revert turns it back into FORMAT_RANGE / SET_COLUMN_WIDTH actions that put
 * the old values back. No new action type: the inverse is the forward action
 * with the old values.
 *
 * Only what the forward action writes is captured, so a snapshot of a bold +
 * fill change holds bold and fill, nothing else.
 */

/** One FORMAT_RANGE call's worth of cells, and the most one snapshot may hold. */
export const FORMAT_SNAPSHOT_MAX_CELLS = 20_000;
/** All FORMAT_RANGE snapshots of one change set together (keeps the apply call and the document small). */
export const FORMAT_SNAPSHOT_MAX_TOTAL_CELLS = 60_000;
export const FORMAT_SNAPSHOT_MAX_COLUMNS = 200;
const MAX_SNAPSHOTS_PER_CHANGE_SET = 500;
const MAX_PALETTE = 5_000;
const MAX_ROW_INDEX = 1_048_575;
const MAX_COL_INDEX = 16_383;

export interface FormatSnapshotCellFormat {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontName?: string;
  fontColor?: string;
  fillColor?: string;
  clearFill?: boolean;
  numberFormat?: string;
  horizontalAlignment?: 'left' | 'center' | 'right' | 'general';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
}

export interface FormatSnapshot {
  kind: 'format' | 'columns';
  sheetName: string;
  /** False when the earlier state could not be read, or not put back exactly. Refuses the revert. */
  restorable: boolean;
  reason?: string;
  /** kind 'format': 0-based origin and size of the range read. */
  row?: number;
  col?: number;
  rowCount?: number;
  colCount?: number;
  /** kind 'format': the distinct cell formats, and for every cell the index of its format. */
  palette?: FormatSnapshotCellFormat[];
  grid?: number[][];
  /** kind 'columns': the width, in points, of each column before it was resized. */
  widths?: { col: number; width: number }[];
}

/** The few fields of an action these checks read, all untyped: a change set's actions are stored as plain JSON. */
type ActionLike = { type: string } & Record<string, unknown>;

const FORMAT_KEYS: Record<keyof FormatSnapshotCellFormat, 'boolean' | 'number' | 'string'> = {
  bold: 'boolean',
  italic: 'boolean',
  fontSize: 'number',
  fontName: 'string',
  fontColor: 'string',
  fillColor: 'string',
  clearFill: 'boolean',
  numberFormat: 'string',
  horizontalAlignment: 'string',
  verticalAlignment: 'string',
  wrapText: 'boolean',
};
const H_ALIGN = new Set(['left', 'center', 'right', 'general']);
const V_ALIGN = new Set(['top', 'middle', 'bottom']);

const isIndex = (value: unknown, max: number): value is number =>
  typeof value === 'number' && Number.isInteger(value) && value >= 0 && value <= max;

function sanitizeCellFormat(raw: unknown): FormatSnapshotCellFormat | null {
  if (!raw || typeof raw !== 'object') return null;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    const expected = FORMAT_KEYS[key as keyof FormatSnapshotCellFormat];
    if (!expected || typeof value !== expected) return null;
    if (typeof value === 'number' && !Number.isFinite(value)) return null;
    if (typeof value === 'string' && value.length > 255) return null;
    if (key === 'horizontalAlignment' && !H_ALIGN.has(value as string)) return null;
    if (key === 'verticalAlignment' && !V_ALIGN.has(value as string)) return null;
    out[key] = value;
  }
  return out as FormatSnapshotCellFormat;
}

function unrestorable(kind: FormatSnapshot['kind'], sheetName: string, reason: string): FormatSnapshot {
  return { kind, sheetName, restorable: false, reason };
}

/**
 * The apply call's body comes from the add-in, and what it holds is later turned
 * into actions that write to the user's workbook. Anything that does not have the
 * expected shape is kept only as a refusal, never as data.
 */
export function sanitizeFormatSnapshots(raw: unknown): FormatSnapshot[] {
  if (!Array.isArray(raw)) return [];
  const out: FormatSnapshot[] = [];
  for (const item of raw.slice(0, MAX_SNAPSHOTS_PER_CHANGE_SET)) {
    const entry = (item ?? {}) as Record<string, unknown>;
    const kind = entry.kind === 'columns' ? 'columns' : entry.kind === 'format' ? 'format' : null;
    const sheetName = typeof entry.sheetName === 'string' ? entry.sheetName.trim() : '';
    if (!kind || !sheetName || sheetName.length > 255) {
      out.push(unrestorable('format', sheetName || '?', 'malformed snapshot'));
      continue;
    }
    if (entry.restorable !== true) {
      const reason = typeof entry.reason === 'string' ? entry.reason.slice(0, 200) : 'not captured';
      out.push(unrestorable(kind, sheetName, reason));
      continue;
    }

    if (kind === 'columns') {
      const widths = Array.isArray(entry.widths) ? entry.widths : null;
      const clean = (widths ?? [])
        .map((w) => w as { col?: unknown; width?: unknown })
        .filter((w) => isIndex(w.col, MAX_COL_INDEX) && typeof w.width === 'number' && Number.isFinite(w.width) && w.width >= 0 && w.width <= 2_000)
        .map((w) => ({ col: w.col as number, width: w.width as number }));
      if (!widths || clean.length !== widths.length || clean.length > FORMAT_SNAPSHOT_MAX_COLUMNS) {
        out.push(unrestorable(kind, sheetName, 'malformed column widths'));
      } else {
        out.push({ kind, sheetName, restorable: true, widths: clean });
      }
      continue;
    }

    const { row, col, rowCount, colCount, palette, grid } = entry as Record<string, unknown>;
    const paletteOk = Array.isArray(palette) && palette.length <= MAX_PALETTE;
    const cleanPalette = paletteOk ? (palette as unknown[]).map(sanitizeCellFormat) : [];
    const shapeOk =
      isIndex(row, MAX_ROW_INDEX) &&
      isIndex(col, MAX_COL_INDEX) &&
      isIndex(rowCount, FORMAT_SNAPSHOT_MAX_CELLS) &&
      isIndex(colCount, FORMAT_SNAPSHOT_MAX_CELLS) &&
      (rowCount as number) * (colCount as number) <= FORMAT_SNAPSHOT_MAX_CELLS;
    const gridOk =
      shapeOk &&
      Array.isArray(grid) &&
      grid.length === rowCount &&
      (grid as unknown[]).every(
        (line) =>
          Array.isArray(line) &&
          line.length === colCount &&
          line.every((cell) => isIndex(cell, Math.max(0, cleanPalette.length - 1))),
      );
    if (!paletteOk || cleanPalette.some((p) => p === null) || !gridOk) {
      out.push(unrestorable(kind, sheetName, 'malformed format snapshot'));
      continue;
    }
    out.push({
      kind,
      sheetName,
      restorable: true,
      row: row as number,
      col: col as number,
      rowCount: rowCount as number,
      colCount: colCount as number,
      palette: cleanPalette as FormatSnapshotCellFormat[],
      grid: grid as number[][],
    });
  }
  return out;
}

// ---- Which actions can be put back ---------------------------------------------------

function letterToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Cells in an A1 range ("A3:D3", "B4"); null for anything else, including whole columns and rows. */
function a1Area(range: string): number | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(range.trim());
  if (!m) return null;
  const c1 = letterToIndex(m[1]);
  const r1 = Number(m[2]);
  const c2 = m[3] ? letterToIndex(m[3]) : c1;
  const r2 = m[4] ? Number(m[4]) : r1;
  return (Math.abs(r2 - r1) + 1) * (Math.abs(c2 - c1) + 1);
}

/** How many cells a FORMAT_RANGE covers, or null when that cannot be told. */
export function formatActionArea(action: ActionLike): number | null {
  if (typeof action.range === 'string' && action.row === undefined) return a1Area(action.range);
  const rows = action.rowCount ?? 1;
  const cols = action.colCount ?? 1;
  if (typeof rows !== 'number' || typeof cols !== 'number') return null;
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1) return null;
  return rows * cols;
}

/**
 * Whether Revert can put this action back. Per instance, like CONDITIONAL_FORMAT:
 * the same action type is restorable for a header band and not for a whole sheet.
 * Borders are never restored: they are per edge, and FORMAT_RANGE cannot say that.
 */
export function isFormatActionRestorable(action: ActionLike): boolean {
  if (action.type === 'FORMAT_RANGE') {
    const format = action.format as { borders?: unknown } | undefined;
    if (!format || typeof format !== 'object' || format.borders) return false;
    const area = formatActionArea(action);
    return area !== null && area <= FORMAT_SNAPSHOT_MAX_CELLS;
  }
  if (action.type === 'AUTOFIT_COLUMNS') {
    if (Array.isArray(action.columns)) return action.columns.length <= FORMAT_SNAPSHOT_MAX_COLUMNS;
    const count = action.colCount;
    return count === undefined || (typeof count === 'number' && Number.isInteger(count) && count <= FORMAT_SNAPSHOT_MAX_COLUMNS);
  }
  return false;
}

/** True when the FORMAT_RANGE actions of one change set add up to more than a snapshot may hold. */
export function formatActionsExceedTotalCap(actions: ActionLike[]): boolean {
  let total = 0;
  for (const action of actions) {
    if (action.type !== 'FORMAT_RANGE') continue;
    total += formatActionArea(action) ?? 0;
  }
  return total > FORMAT_SNAPSHOT_MAX_TOTAL_CELLS;
}

const SNAPSHOT_TYPES = new Set(['FORMAT_RANGE', 'AUTOFIT_COLUMNS']);

/**
 * Fails closed: a change set with restorable formatting actions can only be
 * reverted when a restorable snapshot came back for each of them. Anything less
 * and the revert would report success while leaving the formatting in place.
 */
export function assertFormatSnapshotsComplete(
  changeSetId: string,
  actions: ActionLike[],
  snapshots: FormatSnapshot[],
): void {
  const expected = actions.filter((a) => SNAPSHOT_TYPES.has(a.type) && isFormatActionRestorable(a));
  if (expected.length === 0) return;
  if (snapshots.length < expected.length || snapshots.some((s) => !s.restorable)) {
    const types = [...new Set(expected.map((a) => a.type))];
    throw new RevertNoOpError(changeSetId, types);
  }
}

// ---- Snapshot -> inverse actions ------------------------------------------------------

interface Rect {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
  index: number;
}

/** Splits a grid of palette indices into rectangles of one index each (a row run, extended down while the next row repeats it). */
export function gridToRects(grid: number[][]): Rect[] {
  const rects: Rect[] = [];
  let open = new Map<string, Rect>();
  grid.forEach((line, row) => {
    const next = new Map<string, Rect>();
    let start = 0;
    for (let col = 1; col <= line.length; col += 1) {
      if (col < line.length && line[col] === line[start]) continue;
      const key = `${start}:${col - 1}:${line[start]}`;
      const carried = open.get(key);
      if (carried) {
        carried.rowCount += 1;
        next.set(key, carried);
      } else {
        const rect = { row, col: start, rowCount: 1, colCount: col - start, index: line[start] };
        rects.push(rect);
        next.set(key, rect);
      }
      start = col;
    }
    open = next;
  });
  return rects;
}

/**
 * The actions that put the formatting back, newest snapshot first. Two actions
 * that touched the same cells restore in reverse, so the oldest state lands last.
 */
export function formatSnapshotsToInverseActions(snapshots: FormatSnapshot[]): Action[] {
  const out: Action[] = [];
  for (const snapshot of [...snapshots].reverse()) {
    if (!snapshot.restorable) continue;
    if (snapshot.kind === 'columns') {
      const widths = [...(snapshot.widths ?? [])].sort((a, b) => a.col - b.col);
      let i = 0;
      while (i < widths.length) {
        let j = i;
        while (j + 1 < widths.length && widths[j + 1].col === widths[j].col + 1 && widths[j + 1].width === widths[i].width) j += 1;
        out.push({
          type: 'SET_COLUMN_WIDTH',
          sheetName: snapshot.sheetName,
          col: widths[i].col,
          colCount: j - i + 1,
          width: widths[i].width,
        } as Action);
        i = j + 1;
      }
      continue;
    }
    for (const rect of gridToRects(snapshot.grid ?? [])) {
      const format = snapshot.palette?.[rect.index];
      if (!format || Object.keys(format).length === 0) continue;
      out.push({
        type: 'FORMAT_RANGE',
        sheetName: snapshot.sheetName,
        row: (snapshot.row ?? 0) + rect.row,
        col: (snapshot.col ?? 0) + rect.col,
        rowCount: rect.rowCount,
        colCount: rect.colCount,
        format,
      } as Action);
    }
  }
  return out;
}
