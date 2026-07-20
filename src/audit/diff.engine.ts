import { Action } from '../agents/types/agent.types';
import { ShadowCell, ShadowWorkbook } from '../virtual/shadowWorkbook.types';
import { CellChange, CellSnapshot } from './types/change-set.types';

function parseCellKey(key: string): { sheet: string; address: string } {
  const bang = key.indexOf('!');
  if (bang === -1) {
    return { sheet: '', address: key };
  }
  return { sheet: key.slice(0, bang), address: key.slice(bang + 1) };
}

function cellKey(sheet: string, address: string): string {
  return `${sheet}!${address}`;
}

function snapshotCell(cell: ShadowCell): CellSnapshot {
  const displayValue =
    cell.formula && cell.formula.startsWith('=')
      ? cell.formula
      : cell.value;
  return {
    value: displayValue ?? null,
    formula: cell.formula ?? '',
    format: cell.numberFormat ?? 'General',
  };
}

/** Snapshot all cells in a shadow workbook keyed as "Sheet!A1". */
export function snapshotBeforeState(shadow: ShadowWorkbook): Record<string, CellSnapshot> {
  const state: Record<string, CellSnapshot> = {};
  for (const sheet of shadow.sheets.values()) {
    for (const [address, cell] of sheet.cells) {
      state[cellKey(sheet.name, address)] = snapshotCell(cell);
    }
  }
  return state;
}

function resolveDisplayValue(cell: ShadowCell | undefined): unknown {
  if (!cell) return null;
  if (cell.formula && cell.formula.startsWith('=')) return cell.formula;
  if (cell.formula && cell.formula.startsWith('[formula:')) return cell.formula;
  return cell.value ?? null;
}

function valuesEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (a == null && b == null) return true;
  return String(a ?? '') === String(b ?? '');
}

function isHardcodedValue(cell: ShadowCell | undefined, formula: string): boolean {
  if (!cell) return false;
  if (formula && formula.startsWith('=')) return false;
  return cell.value !== null && cell.value !== undefined && String(cell.value).trim() !== '';
}

/** Diff before vs after shadow workbooks; uses changedCells when available. */
export function generateDiff(
  before: ShadowWorkbook,
  after: ShadowWorkbook,
): CellChange[] {
  const changes: CellChange[] = [];
  const keys =
    after.changedCells.size > 0
      ? after.changedCells
      : collectAllCellKeys(before, after);

  for (const key of keys) {
    const { sheet, address } = parseCellKey(key);
    const beforeCell = before.sheets.get(sheet)?.cells.get(address);
    const afterCell = after.sheets.get(sheet)?.cells.get(address);
    const beforeVal = resolveDisplayValue(beforeCell);
    const afterVal = resolveDisplayValue(afterCell);

    if (valuesEqual(beforeVal, afterVal)) continue;

    const formula = afterCell?.formula && afterCell.formula.startsWith('=')
      ? afterCell.formula
      : undefined;

    changes.push({
      cell: address,
      sheet,
      before: beforeVal ?? null,
      after: afterVal ?? null,
      formula,
      isHardcoded: isHardcodedValue(afterCell, afterCell?.formula ?? ''),
    });
  }

  return changes.sort((a, b) =>
    `${a.sheet}!${a.cell}`.localeCompare(`${b.sheet}!${b.cell}`, undefined, { numeric: true }),
  );
}

function collectAllCellKeys(before: ShadowWorkbook, after: ShadowWorkbook): Set<string> {
  const keys = new Set<string>();
  for (const sheet of before.sheets.values()) {
    for (const address of sheet.cells.keys()) {
      keys.add(cellKey(sheet.name, address));
    }
  }
  for (const sheet of after.sheets.values()) {
    for (const address of sheet.cells.keys()) {
      keys.add(cellKey(sheet.name, address));
    }
  }
  return keys;
}

/** Build inverse SET_CELL / SET_FORMULA actions to restore beforeState. */
export function beforeStateToInverseActions(
  beforeState: Record<string, CellSnapshot>,
  changes: CellChange[],
): Action[] {
  const actions: Action[] = [];
  const touched = new Set(changes.map((c) => cellKey(c.sheet, c.cell)));

  for (const key of touched) {
    const snapshot = beforeState[key];
    const { sheet, address } = parseCellKey(key);
    if (!snapshot) {
      actions.push({
        type: 'SET_CELL',
        sheetName: sheet,
        address,
        value: null,
        explicitOverwriteConfirmed: true,
      } as Action);
      continue;
    }

    if (snapshot.formula && snapshot.formula.startsWith('=')) {
      actions.push({
        type: 'SET_FORMULA',
        sheetName: sheet,
        address,
        formula: snapshot.formula,
        explicitOverwriteConfirmed: true,
      } as Action);
    } else {
      actions.push({
        type: 'SET_CELL',
        sheetName: sheet,
        address,
        value: snapshot.value as string | number | boolean | null,
        explicitOverwriteConfirmed: true,
      } as Action);
    }
  }

  return actions;
}
