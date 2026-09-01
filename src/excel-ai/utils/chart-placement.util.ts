import { SheetActionPayload } from '../types/sheet-actions.types';

/**
 * Keep a chart off its own source data — TASKS.md #133 / #163.
 *
 * A live build placed a chart at `D4:K18` over a source range of `A4:D16`, so
 * the chart covered column D — the "Pending Amount" column of the very table it
 * was plotting. The planner prescribes that anchor literally, and the Executor
 * transcribes it faithfully, so neither is "wrong"; the anchor itself is a
 * stale constant. `executor.prompt.ts`'s CREATE_CHART example pairs
 * `startCell D4` with a **2-column** source (`A4:B9`), where D genuinely is the
 * first free column. The ledger rule reused `D4` against a **4-column** table
 * without re-deriving it.
 *
 * Fixed in code rather than only in the prompt, for the same reason the
 * presentation and consolidation passes are: a rule the model must remember is
 * a rule it will sometimes forget, and this one is silently destructive —
 * nothing errors, a column of the user's dashboard is just hidden. The prompt
 * is corrected too, so the model stops emitting it; this guarantees the outcome
 * either way.
 *
 * Deliberately minimal: it only ever moves a chart RIGHT, never resizes it,
 * never touches a chart already clear of its source, and never moves cell
 * content. Layout-preserving, exactly like the presentation pass.
 */

/** Blank columns left between a source table and the chart beside it. */
const GUTTER_COLUMNS = 1;

export function columnLetterToIndex(letters: string): number {
  let index = 0;
  for (const ch of letters.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64);
  return index - 1;
}

export function columnIndexToLetter(col: number): string {
  let index = col + 1;
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

interface CellRef {
  col: number;
  row: number;
}

export function parseCellRef(address: string): CellRef | null {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(address ?? '').trim());
  if (!m) return null;
  return { col: columnLetterToIndex(m[1]), row: Number(m[2]) - 1 };
}

/** Last column index of a range like `A4:D16`, or of a bare cell. */
export function lastColumnOfRange(range: string): number | null {
  const local = String(range ?? '').includes('!')
    ? String(range).slice(String(range).indexOf('!') + 1)
    : String(range ?? '');
  const parts = local.split(':');
  const end = parseCellRef(parts[parts.length - 1]);
  const start = parseCellRef(parts[0]);
  if (!end && !start) return null;
  return Math.max(end?.col ?? -1, start?.col ?? -1);
}

/**
 * Move any chart whose anchor falls inside its own source columns to the first
 * free column to the right, preserving its width.
 */
export function applyChartPlacementPass(actions: SheetActionPayload[]): SheetActionPayload[] {
  let changed = false;

  const next = actions.map((action) => {
    if (action.type !== 'CREATE_CHART') return action;

    const sourceRange = String(action.sourceRange ?? '');
    const start = parseCellRef(String(action.startCell ?? ''));
    const sourceLastCol = lastColumnOfRange(sourceRange);
    if (!start || sourceLastCol === null) return action;

    // A chart on a DIFFERENT sheet from its source cannot overlap it.
    const sourceSheet = String(action.sourceSheetName ?? action.sheetName ?? '').trim().toLowerCase();
    const chartSheet = String(action.sheetName ?? '').trim().toLowerCase();
    if (sourceSheet && chartSheet && sourceSheet !== chartSheet) return action;

    const firstFreeCol = sourceLastCol + 1 + GUTTER_COLUMNS;
    if (start.col >= firstFreeCol) return action;

    const shift = firstFreeCol - start.col;
    const end = parseCellRef(String(action.endCell ?? ''));

    changed = true;
    return {
      ...action,
      startCell: `${columnIndexToLetter(start.col + shift)}${start.row + 1}`,
      ...(end
        ? { endCell: `${columnIndexToLetter(end.col + shift)}${end.row + 1}` }
        : {}),
    };
  });

  return changed ? next : actions;
}
