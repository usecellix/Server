import { parseA1Cell } from './range-merge.util';

/**
 * Collapse identical value writes within one wave — TASKS.md #317.
 *
 * Parallel subtasks each see the same prior workbook, so when several decide a
 * shared sheet needs the same content, each writes it. A live run had eleven
 * month subtasks each emit the same `BATCH_SET` onto Lists (D1 = "Unit No",
 * ...): the first copy landed, the client's overwrite guard refused the second
 * as an overwrite of an occupied cell, and the step could never be accepted.
 *
 * A write that sets the SAME cell to the SAME value/formula as an earlier write
 * in the batch changes nothing, so it is dropped. Two DIFFERENT values for one
 * cell are a real conflict and are left alone — silently picking one would hide
 * it; `conflicts` reports them instead.
 */

interface WriteLike {
  type?: string;
  sheetName?: string;
  row?: number;
  col?: number;
  address?: string;
  value?: unknown;
  formula?: unknown;
  operations?: Array<Record<string, unknown>>;
}

function cellKey(sheet: unknown, cell: { row?: unknown; col?: unknown; address?: unknown }): string | null {
  let row: number | undefined;
  let col: number | undefined;
  if (typeof cell.row === 'number' && typeof cell.col === 'number') {
    row = cell.row;
    col = cell.col;
  } else if (typeof cell.address === 'string') {
    const parsed = parseA1Cell(cell.address.replace(/\$/g, ''));
    if (parsed) ({ row, col } = parsed);
  }
  if (row === undefined || col === undefined) return null;
  return `${String(sheet ?? '').toLowerCase()}|${row}|${col}`;
}

function contentOf(write: { value?: unknown; formula?: unknown }): string {
  return JSON.stringify([write.formula ?? null, write.value ?? null]);
}

export function dedupeIdenticalWrites<T>(actions: T[]): {
  actions: T[];
  removed: number;
  conflicts: string[];
} {
  const seen = new Map<string, string>();
  const conflicts = new Set<string>();
  let removed = 0;
  const out: T[] = [];

  /** true → keep this write. */
  const admit = (key: string | null, content: string): boolean => {
    if (!key) return true;
    const prior = seen.get(key);
    if (prior === undefined) {
      seen.set(key, content);
      return true;
    }
    if (prior === content) {
      removed += 1;
      return false;
    }
    conflicts.add(key);
    return true;
  };

  for (const action of actions) {
    const write = action as unknown as WriteLike;

    if (write.type === 'SET_CELL' || write.type === 'SET_FORMULA') {
      if (admit(cellKey(write.sheetName, write), contentOf(write))) out.push(action);
      continue;
    }

    if (write.type === 'BATCH_SET' && Array.isArray(write.operations)) {
      const operations = write.operations.filter((op) =>
        admit(
          cellKey(op.sheetName ?? write.sheetName, op),
          contentOf(op as { value?: unknown; formula?: unknown }),
        ),
      );
      if (operations.length === write.operations.length) out.push(action);
      else if (operations.length > 0) out.push({ ...(action as object), operations } as T);
      continue;
    }

    out.push(action);
  }

  return { actions: out, removed, conflicts: [...conflicts] };
}
