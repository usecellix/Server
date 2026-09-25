import { Action } from '../types/agent.types';
import { parseA1Cell, columnLetterToIndex } from './range-merge.util';

/**
 * Point every list dropdown at the list it is named for — TASKS.md #334.
 *
 * The Lists sheet and the month sheets' dropdowns are planned by different
 * subtasks, and the month step GUESSES the Lists layout. Live
 * run_1790332994308_etr7oxl: Lists had Source in A, Payment Status in B and
 * Bank Account in C, values from row 2; the month dropdowns pointed at
 * `Lists!$B$3:$B$10` (Payment Status, minus "Paid"), `Lists!$C$3:$C$12` (the
 * BANK ACCOUNTS, under Source) and `Lists!$D$3:$D$6` (an empty column, under
 * Bank Account). The user saw "Account 2" in the source column.
 *
 * Both layouts are in the run's own actions, so this is not a guess: the
 * dropdown's column header on its sheet names the list, the list sheet's row-1
 * header says which column holds it, and its filled cells say how far it goes.
 * A dropdown whose name matches no list header is left exactly as it was.
 */

type Grid = Map<string, Map<string, unknown>>; // sheet -> "A1" -> value

const norm = (value: unknown) => String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');

function letter(zeroBased: number): string {
  let n = zeroBased + 1;
  let out = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    out = String.fromCharCode(65 + mod) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** Every literal value these actions write, by sheet and cell. */
function writtenGrid(actions: Action[]): Grid {
  const grid: Grid = new Map();
  const put = (sheet: unknown, address: string, value: unknown) => {
    const key = norm(sheet);
    if (!key || value === undefined) return;
    if (!grid.has(key)) grid.set(key, new Map());
    grid.get(key)!.set(address.toUpperCase().replace(/\$/g, ''), value);
  };
  for (const action of actions) {
    const a = action as unknown as Record<string, unknown>;
    if (a.type === 'BATCH_SET' && Array.isArray(a.operations)) {
      for (const op of a.operations as Array<Record<string, unknown>>) {
        const address =
          typeof op.address === 'string'
            ? op.address
            : typeof op.row === 'number' && typeof op.col === 'number'
              ? `${letter(op.col)}${op.row + 1}`
              : null;
        if (address && op.formula === undefined) put(op.sheetName ?? a.sheetName, address, op.value);
      }
    } else if (a.type === 'SET_CELL' && typeof a.address === 'string') {
      put(a.sheetName, a.address, a.value);
    }
  }
  return grid;
}

const LIST_SOURCE = /^=?'?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i;

export interface ListRepair {
  sheet: string;
  range: string;
  from: string;
  to: string;
}

/**
 * Rewrites `listSource` on DATA_VALIDATION actions in `actions` IN PLACE, using
 * `context` (every action the run has produced so far, these included) as the
 * source of both layouts. Returns what it changed, for the log.
 */
export function repairListValidationSources(actions: Action[], context: Action[]): ListRepair[] {
  const grid = writtenGrid(context);
  const repairs: ListRepair[] = [];

  for (const action of actions) {
    const a = action as unknown as {
      type: string;
      sheetName?: string;
      range?: string;
      validation?: { kind?: string; listSource?: string; promptTitle?: string };
    };
    if (a.type !== 'DATA_VALIDATION' || a.validation?.kind !== 'list') continue;
    const source = a.validation.listSource ?? '';
    const parsed = LIST_SOURCE.exec(source.trim());
    if (!parsed || !a.range) continue;
    const [, listSheetRaw, fromCol, fromRow] = parsed;
    const listSheet = grid.get(norm(listSheetRaw));
    if (!listSheet) continue;

    // The list this dropdown is FOR: its own column's header, else its title.
    const target = parseA1Cell(a.range.split(':')[0].replace(/\$/g, ''));
    const ownHeader = target ? grid.get(norm(a.sheetName))?.get(`${letter(target.col)}1`) : undefined;
    const wanted = norm(ownHeader ?? a.validation.promptTitle);
    if (!wanted) continue;

    // The list sheet's column carrying that header in row 1.
    let listCol: number | null = null;
    for (const [address, value] of listSheet) {
      const cell = parseA1Cell(address);
      if (cell && cell.row === 0 && norm(value) === wanted) listCol = cell.col;
    }
    if (listCol === null) continue;

    let lastRow = 1;
    for (const [address, value] of listSheet) {
      const cell = parseA1Cell(address);
      if (cell && cell.col === listCol && cell.row > 0 && norm(value) !== '') {
        lastRow = Math.max(lastRow, cell.row + 1);
      }
    }
    if (lastRow < 2) continue;

    const col = letter(listCol);
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(listSheetRaw) ? listSheetRaw : `'${listSheetRaw.replace(/'/g, "''")}'`;
    const repaired = `${quoted}!$${col}$2:$${col}$${lastRow}`;
    const alreadyRight =
      columnLetterToIndex(fromCol) === listCol && Number(fromRow) === 2 && Number(parsed[5] ?? fromRow) === lastRow;
    if (alreadyRight) continue;

    a.validation.listSource = repaired;
    repairs.push({ sheet: a.sheetName ?? '', range: a.range, from: source, to: repaired });
  }
  return repairs;
}
