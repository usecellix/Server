import { SheetActionPayload } from '../types/sheet-actions.types';
import { WorkbookContext } from '../../types/cellix.types';

/**
 * A request to delete rows BY CONDITION ("delete blank rows", "delete rows
 * where GSTIN is blank") whose row numbers the model invented.
 *
 * Live runs against a sheet with no blank rows at all produced
 * `DELETE_ROW row:10 rowCount:21` and, on a re-run, `DELETE_ROW row:1
 * rowCount:30` — 21 and then all 30 rows of real data — each reported as
 * "verified: true", because no checker compared the targeted rows against
 * their actual contents. DELETE_MATCHING_ROWS exists so the predicate is
 * resolved against real cells at apply time; this guard is the net for when
 * the model reaches for DELETE_ROW anyway.
 *
 * Deliberately fails closed: if the condition is "blank" and the targeted rows
 * cannot be read from context, the action is dropped rather than trusted.
 * TASKS.md #234.
 */

/** The user described WHICH rows rather than naming them. */
const CONDITIONAL_ROW_PHRASE =
  /\b(blank|empty|duplicate|duplicates|where|matching|that\s+(?:have|has|are|contain)|with\s+no|without|missing)\b/i;

/** An explicit row number ("delete row 7", "rows 10-12") — the user named the rows themselves. */
const EXPLICIT_ROW_NUMBER = /\brows?\s+\d+/i;

/** The condition is specifically emptiness, which is checkable against the data. */
const BLANKNESS_PHRASE = /\b(blank|empty|with\s+no\s+data|missing)\b/i;

export interface ConditionalRowDeleteGuardResult {
  actions: SheetActionPayload[];
  /** Human-readable reasons for anything dropped, for logging. */
  dropped: string[];
}

function sheetRowsFrom(
  context: WorkbookContext | undefined,
  sheetName: string | undefined,
): unknown[][] | null {
  if (!context?.sheets?.length) return null;
  const target =
    context.sheets.find((sheet: { sheetName?: string }) => sheet.sheetName === sheetName) ??
    context.sheets.find((sheet: { sheetName?: string }) => sheet.sheetName === context.activeSheet);
  const rows = target?.sampleData;
  return Array.isArray(rows) && rows.length > 0 ? (rows as unknown[][]) : null;
}

function cellIsBlank(cell: unknown): boolean {
  return cell === null || cell === undefined || String(cell).trim() === '';
}

function rowIsBlank(row: unknown[] | undefined): boolean {
  if (!row) return false;
  return row.every(cellIsBlank);
}

/**
 * "Delete rows where column D is blank" scopes the blankness to ONE column, so
 * demanding the whole row be empty would block a correct delete — which it did
 * on the first run of this guard. Resolve the named column (letter or header)
 * so the check matches what the user actually asked.
 */
function resolveConditionColumn(
  message: string,
  headers: string[] | undefined,
): number | null {
  const letter = /\bcolumn\s+([A-Z])\b/i.exec(message)?.[1];
  if (letter) return letter.toUpperCase().charCodeAt(0) - 65;

  if (!headers?.length) return null;
  const byName = headers.findIndex(
    (header) => header && new RegExp(`\\b${header.trim()}\\b`, 'i').test(message),
  );
  return byName >= 0 ? byName : null;
}

function headersFrom(
  context: WorkbookContext | undefined,
  sheetName: string | undefined,
): string[] | undefined {
  if (!context?.sheets?.length) return undefined;
  const target =
    context.sheets.find((sheet: { sheetName?: string }) => sheet.sheetName === sheetName) ??
    context.sheets.find((sheet: { sheetName?: string }) => sheet.sheetName === context.activeSheet);
  return target?.headers;
}

export function guardConditionalRowDeletes(
  actions: SheetActionPayload[],
  userMessage: string | undefined,
  context?: WorkbookContext,
): ConditionalRowDeleteGuardResult {
  const message = String(userMessage ?? '');
  const conditional = CONDITIONAL_ROW_PHRASE.test(message) && !EXPLICIT_ROW_NUMBER.test(message);
  if (!conditional) return { actions, dropped: [] };

  const dropped: string[] = [];
  const kept = actions.filter((action) => {
    if (action.type !== 'DELETE_ROW') return true;

    const startRow = typeof action.row === 'number' ? action.row : null;
    if (startRow === null) {
      dropped.push('DELETE_ROW without a row index on a conditional delete');
      return false;
    }
    const rowCount = Math.max(1, Number(action.rowCount ?? 1));

    // Only blankness is checkable from context; any other condition means the
    // model picked rows by reasoning, which is exactly what must not be trusted.
    if (!BLANKNESS_PHRASE.test(message)) {
      dropped.push(
        `DELETE_ROW row=${startRow} rowCount=${rowCount} for a conditional delete — use DELETE_MATCHING_ROWS`,
      );
      return false;
    }

    const rows = sheetRowsFrom(context, action.sheetName);
    if (!rows) {
      dropped.push(
        `DELETE_ROW row=${startRow} rowCount=${rowCount} could not be checked against sheet data`,
      );
      return false;
    }

    const headers = headersFrom(context, action.sheetName);
    const conditionColumn = resolveConditionColumn(message, headers);
    const isEmptyEnough = (row: unknown[]): boolean =>
      conditionColumn === null ? rowIsBlank(row) : cellIsBlank(row[conditionColumn]);

    // sampleData excludes the header row, so sheet row N is sampleData[N - 1]
    // when a header is present; the executor is inconsistent about which index
    // it means, so accept either and drop only when neither reading is blank.
    for (let offset = 0; offset < rowCount; offset += 1) {
      const sheetRowIndex = startRow + offset;
      const candidates = [rows[sheetRowIndex], rows[sheetRowIndex - 1]].filter(
        (row): row is unknown[] => Array.isArray(row),
      );
      if (candidates.length === 0 || !candidates.some(isEmptyEnough)) {
        dropped.push(
          `DELETE_ROW row=${startRow} rowCount=${rowCount} targets non-empty row ${sheetRowIndex}`,
        );
        return false;
      }
    }

    return true;
  });

  return { actions: kept, dropped };
}
