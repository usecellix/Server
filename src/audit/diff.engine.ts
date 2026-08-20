import { Action } from '../agents/types/agent.types';
import { EXCEL_ERRORS } from '../formula/formula-validator.service';
import { colIndexToLetter, letterToColIndex } from '../virtual/shadowWorkbook';
import { ShadowCell, ShadowSheet, ShadowWorkbook } from '../virtual/shadowWorkbook.types';
import { CellChange, CellSnapshot, FormulaErrorChange, StructuralOp } from './types/change-set.types';

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
    bold: cell.bold,
    italic: cell.italic,
    fontColor: cell.fontColor,
    fillColor: cell.fillColor,
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

/**
 * Inverse of `snapshotBeforeState`: rebuild a `ShadowWorkbook` from a stored
 * `ChangeSet.beforeState` map, so revert-time verification (TASKS.md #19) can replay the
 * original forward `actions` and compare against a real shadow, without needing to have
 * persisted the full post-apply shadow separately.
 */
export function shadowFromBeforeState(beforeState: Record<string, CellSnapshot>): ShadowWorkbook {
  const sheets = new Map<string, ShadowSheet>();

  for (const [key, snapshot] of Object.entries(beforeState)) {
    const { sheet: sheetName, address } = parseCellKey(key);
    let sheet = sheets.get(sheetName);
    if (!sheet) {
      sheet = { name: sheetName, cells: new Map(), rowCount: 0, columnCount: 0, structure: 'unknown' };
      sheets.set(sheetName, sheet);
    }
    sheet.cells.set(address, {
      value: snapshot.value ?? null,
      formula: snapshot.formula && snapshot.formula.startsWith('=') ? snapshot.formula : '',
      numberFormat: snapshot.format ?? 'General',
      bold: snapshot.bold,
      italic: snapshot.italic,
      fontColor: snapshot.fontColor,
      fillColor: snapshot.fillColor,
    });
    const row = Number.parseInt(address.replace(/[A-Z]+/i, ''), 10);
    const col = letterToColIndex(address.replace(/\d+/g, ''));
    sheet.rowCount = Math.max(sheet.rowCount, row);
    sheet.columnCount = Math.max(sheet.columnCount, col + 1);
  }

  return {
    activeSheetName: sheets.keys().next().value ?? '',
    sheets,
    namedRanges: new Map(),
    tables: [],
    changedCells: new Set(),
    structuralLog: [],
    conditionalFormats: [],
  };
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

/**
 * PRD A5 ("unintended modification rate", TASKS.md #48) — flag any applied cell change on a
 * sheet no action in this batch declared touching. Deliberately sheet-granularity, not
 * cell/range-granularity: `SheetActionPayload.sheetName` is the only machine-readable
 * "declared target" available today — the Planner's own `SubTask` type carries a
 * `targetSheet` but nothing finer, and inventing a cell-level declared-scope field is out
 * of this task's scope. This still catches the real failure this metric exists for (an
 * action mis-resolving its range and mutating a sheet nobody asked it to touch), just not
 * a same-sheet-wrong-range variant of it.
 *
 * Fails open, not closed: if no action in the batch declares a `sheetName` at all (some
 * action shapes legitimately don't carry one), the scope is unknowable — return no
 * violations rather than flag every change, since a false "everything is unintended" is
 * worse noise than an unmeasured batch.
 */
export function computeUnintendedChanges(changes: CellChange[], actions: Action[]): CellChange[] {
  const declaredSheets = new Set(
    actions.map((action) => action.sheetName).filter((name): name is string => Boolean(name)),
  );
  if (declaredSheets.size === 0) return [];
  return changes.filter((change) => !declaredSheets.has(change.sheet));
}

function findExcelError(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  return EXCEL_ERRORS.find((err) => value.includes(err));
}

/**
 * PRD A6 ("formula error rate", TASKS.md #49) — an Excel error string present in a cell's
 * *after* value that wasn't already there (as the same error) in its *before* value.
 * Deliberately scoped to `changes` (cells this change set actually touched), not a
 * whole-shadow scan: that's what makes "a pre-existing unrelated error elsewhere isn't
 * double-counted" true for free — an error sitting in a cell nobody touched never appears
 * in `changes` at all, so it's never a candidate.
 */
export function detectIntroducedFormulaErrors(changes: CellChange[]): FormulaErrorChange[] {
  const introduced: FormulaErrorChange[] = [];
  for (const change of changes) {
    const afterError = findExcelError(change.after);
    if (!afterError) continue;
    const beforeError = findExcelError(change.before);
    if (beforeError === afterError) continue;
    introduced.push({ cell: change.cell, sheet: change.sheet, error: afterError });
  }
  return introduced;
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

/**
 * Build the `format` object an inverse SET_CELL/SET_FORMULA action attaches to restore
 * a cell's captured formatting. Only includes keys that were actually captured —
 * `applyFormat`/`cell.handler.ts` only touch `FormatSpec` keys that are present, so an
 * omitted key (bold/italic/fontColor/fillColor absent because the source `SheetContext`
 * predates TASKS.md #64, or numberFormat/format simply wasn't captured) leaves that
 * property untouched on revert rather than clobbering it with a wrong value.
 */
function buildRestoreFormat(snapshot: {
  format?: string;
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  fillColor?: string;
}): Record<string, unknown> {
  return {
    ...(snapshot.format !== undefined ? { numberFormat: snapshot.format } : {}),
    ...(snapshot.bold !== undefined ? { bold: snapshot.bold } : {}),
    ...(snapshot.italic !== undefined ? { italic: snapshot.italic } : {}),
    ...(snapshot.fontColor !== undefined ? { fontColor: snapshot.fontColor } : {}),
    ...(snapshot.fillColor !== undefined ? { fillColor: snapshot.fillColor } : {}),
  };
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

    // Restore the captured formatting alongside the value/formula (TASKS.md #10, #64).
    const format = buildRestoreFormat(snapshot);

    if (snapshot.formula && snapshot.formula.startsWith('=')) {
      actions.push({
        type: 'SET_FORMULA',
        sheetName: sheet,
        address,
        formula: snapshot.formula,
        format,
        explicitOverwriteConfirmed: true,
      } as Action);
    } else {
      actions.push({
        type: 'SET_CELL',
        sheetName: sheet,
        address,
        value: snapshot.value as string | number | boolean | null,
        format,
        explicitOverwriteConfirmed: true,
      } as Action);
    }
  }

  return actions;
}

interface StructuralCellSnapshot {
  address: string;
  value: unknown;
  formula: string;
  format: string;
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  fillColor?: string;
}

/**
 * MERGE_CELLS/UNMERGE_CELLS's canonical wire shape carries `range: string`, but
 * virtualApply.ts's own `resolveSpan()` also accepts a `row`/`col`/`rowCount`/`colCount`
 * shape. Normalize to an A1 string here (not reusing `resolveSpan`, which lives in
 * `agents/utils` and depends on `parseA1Range` for the string case — this only needs
 * the reverse direction: build a string, not parse one).
 */
function resolveRangeString(action: Action): string | undefined {
  if (typeof action.range === 'string' && action.range) return action.range;
  if (typeof action.row === 'number' && typeof action.col === 'number') {
    const rowCount =
      typeof action.rowCount === 'number' && action.rowCount > 0 ? action.rowCount : 1;
    const colCount =
      typeof action.colCount === 'number' && action.colCount > 0 ? action.colCount : 1;
    const startAddr = `${colIndexToLetter(action.col)}${action.row + 1}`;
    const endAddr = `${colIndexToLetter(action.col + colCount - 1)}${action.row + rowCount}`;
    return `${startAddr}:${endAddr}`;
  }
  return undefined;
}

function toStructuralCellSnapshot(address: string, cell: ShadowCell): StructuralCellSnapshot {
  return {
    address,
    value: cell.value ?? null,
    formula: cell.formula ?? '',
    format: cell.numberFormat ?? 'General',
    bold: cell.bold,
    italic: cell.italic,
    fontColor: cell.fontColor,
    fillColor: cell.fillColor,
  };
}

function snapshotSheetCells(sheet: ShadowSheet): StructuralCellSnapshot[] {
  const cells: StructuralCellSnapshot[] = [];
  for (const [address, cell] of sheet.cells) {
    cells.push(toStructuralCellSnapshot(address, cell));
  }
  return cells;
}

/** Snapshot only the cells in one 0-based column index, e.g. for a DELETE_COLUMN's own inverse. */
function snapshotColumnCells(sheet: ShadowSheet, colIndex: number): StructuralCellSnapshot[] {
  const cells: StructuralCellSnapshot[] = [];
  for (const [address, cell] of sheet.cells) {
    const col = letterToColIndex(address.replace(/\d+/g, ''));
    if (col !== colIndex) continue;
    cells.push(toStructuralCellSnapshot(address, cell));
  }
  return cells;
}

/** Snapshot only the cells in one 1-based Excel row number, e.g. for a DELETE_ROW's own inverse. */
function snapshotRowCells(sheet: ShadowSheet, rowNumber: number): StructuralCellSnapshot[] {
  const cells: StructuralCellSnapshot[] = [];
  for (const [address, cell] of sheet.cells) {
    const row = Number.parseInt(address.replace(/[A-Z]+/i, ''), 10);
    if (row !== rowNumber) continue;
    cells.push(toStructuralCellSnapshot(address, cell));
  }
  return cells;
}

/**
 * Capture structural (non-cell-shaped) operations from this batch of actions, using the
 * *before* shadow workbook so a DELETE_SHEET's full sheet contents (or a DELETE_COLUMN's
 * deleted column) are available for their inverse — independent of `changes`/`beforeState`'s
 * cell-diff machinery (see TASKS.md #12), which can silently omit them when other actions in
 * the same batch dominate `changedCells`, or — for column shifts — would otherwise conflict
 * with the structural shift the inverse action itself performs (see TASKS.md #13).
 *
 * Column ops are read from `afterShadow.structuralLog` (populated by virtualApply as it
 * *resolves* each action's targeting, e.g. a header name or letter into a column index)
 * rather than re-parsed from `actions` here, so this never drifts from what virtualApply
 * actually did.
 */
export function captureStructuralOps(
  beforeShadow: ShadowWorkbook,
  afterShadow: ShadowWorkbook,
  actions: Action[],
): StructuralOp[] {
  const ops: StructuralOp[] = [];

  for (const action of actions) {
    if (action.type === 'ADD_SHEET' || action.type === 'CREATE_SHEET') {
      const sheetName = action.sheetName ?? action.name;
      if (!sheetName) continue;
      ops.push({
        opType: 'ADD_SHEET',
        sheetName,
        params: {},
        appliedAt: new Date(),
      });
      continue;
    }

    if (action.type === 'DELETE_SHEET') {
      const sheetName = action.sheetName;
      if (!sheetName) continue;
      const sheet = beforeShadow.sheets.get(sheetName);
      ops.push({
        opType: 'DELETE_SHEET',
        sheetName,
        params: {
          rowCount: sheet?.rowCount ?? 0,
          columnCount: sheet?.columnCount ?? 0,
          cells: sheet ? snapshotSheetCells(sheet) : [],
        },
        appliedAt: new Date(),
      });
      continue;
    }

    if (action.type === 'CREATE_TABLE') {
      const sheetName = action.sheetName;
      const tableName = action.tableName;
      if (!sheetName || !tableName) continue;
      ops.push({
        opType: 'CREATE_TABLE',
        sheetName,
        params: { tableName },
        appliedAt: new Date(),
      });
      continue;
    }

    if (action.type === 'MERGE_CELLS' || action.type === 'UNMERGE_CELLS') {
      const sheetName = action.sheetName;
      const range = resolveRangeString(action);
      if (!sheetName || !range) continue;
      ops.push({
        opType: action.type,
        sheetName,
        params: { range },
        appliedAt: new Date(),
      });
      continue;
    }

    // TASKS.md #40 — only a create (no existingRuleId) needs a revert entry; a
    // MODIFY (existingRuleId set) mutates an existing rule in place and isn't
    // captured here at all — see reversibility-catalog.ts's per-instance note.
    // `ruleId` is deliberately absent from `params` at this point: Office.js
    // only assigns it once the real Excel apply happens, well after preview.
    // `ChangeSetService.markApplied()` patches it in before the change set is
    // marked applied — structuralOpsToInverseActions() refuses to build an
    // inverse if that patch never happened, rather than silently no-op.
    if (action.type === 'CONDITIONAL_FORMAT' && !action.existingRuleId) {
      const sheetName = action.sheetName;
      const range = action.range;
      if (!sheetName || !range) continue;
      ops.push({
        opType: 'CONDITIONAL_FORMAT',
        sheetName,
        params: { range },
        appliedAt: new Date(),
      });
      continue;
    }

    // TASKS.md #15 — `chartId` is deliberately absent from `params` at this point:
    // Office.js only assigns the chart's real name once the real Excel apply happens
    // (`handleCreateChart`'s post-create read-back of `chart.name`), well after preview.
    // `ChangeSetService.markApplied()` patches it in, matched by sheetName+sourceRange
    // (the same correlation-key approach #40 used for CONDITIONAL_FORMAT's range),
    // before the change set is marked applied — structuralOpsToInverseActions() refuses
    // to build an inverse if that patch never happened, rather than silently no-op.
    if (action.type === 'CREATE_CHART') {
      const sheetName = action.sheetName;
      const sourceRange = action.sourceRange;
      if (!sheetName || !sourceRange) continue;
      ops.push({
        opType: 'CREATE_CHART',
        sheetName,
        params: { sourceRange },
        appliedAt: new Date(),
      });
    }
  }

  for (const entry of afterShadow.structuralLog) {
    if (entry.opType === 'INSERT_COLUMN') {
      ops.push({
        opType: 'INSERT_COLUMN',
        sheetName: entry.sheetName,
        params: { index: entry.index, count: entry.count },
        appliedAt: new Date(),
      });
      continue;
    }

    if (entry.opType === 'DELETE_COLUMN') {
      const sheet = beforeShadow.sheets.get(entry.sheetName);
      ops.push({
        opType: 'DELETE_COLUMN',
        sheetName: entry.sheetName,
        params: {
          index: entry.index,
          count: entry.count,
          cells: sheet ? snapshotColumnCells(sheet, entry.index) : [],
        },
        appliedAt: new Date(),
      });
      continue;
    }

    if (entry.opType === 'INSERT_ROW') {
      ops.push({
        opType: 'INSERT_ROW',
        sheetName: entry.sheetName,
        params: { index: entry.index, count: entry.count },
        appliedAt: new Date(),
      });
      continue;
    }

    if (entry.opType === 'DELETE_ROW') {
      const sheet = beforeShadow.sheets.get(entry.sheetName);
      ops.push({
        opType: 'DELETE_ROW',
        sheetName: entry.sheetName,
        params: {
          index: entry.index,
          count: entry.count,
          cells: sheet ? snapshotRowCells(sheet, entry.index) : [],
        },
        appliedAt: new Date(),
      });
    }
  }

  return ops;
}

/**
 * Cell-level changes that a structural column/row op's own inverse already fully accounts
 * for (the shift it performs), and which the generic cell-diff inverse must NOT also touch —
 * doing so would either duplicate work harmlessly or, worse, re-create a phantom cell/column/
 * row at an address the structural inverse already removed. Filters to sheet + column/row at
 * or past the lowest structurally-touched index in that sheet (see TASKS.md #13 for the
 * worked column example; rows follow the same shape).
 */
export function excludeStructurallyOwnedChanges(
  changes: CellChange[],
  structuralOps: StructuralOp[],
): CellChange[] {
  const minColBySheet = new Map<string, number>();
  const minRowBySheet = new Map<string, number>();
  for (const op of structuralOps) {
    const index = op.params.index as number | undefined;
    if (index === undefined) continue;
    if (op.opType === 'INSERT_COLUMN' || op.opType === 'DELETE_COLUMN') {
      const current = minColBySheet.get(op.sheetName);
      if (current === undefined || index < current) minColBySheet.set(op.sheetName, index);
    } else if (op.opType === 'INSERT_ROW' || op.opType === 'DELETE_ROW') {
      const current = minRowBySheet.get(op.sheetName);
      if (current === undefined || index < current) minRowBySheet.set(op.sheetName, index);
    }
  }
  if (minColBySheet.size === 0 && minRowBySheet.size === 0) return changes;

  return changes.filter((change) => {
    const minCol = minColBySheet.get(change.sheet);
    const minRow = minRowBySheet.get(change.sheet);
    if (minCol === undefined && minRow === undefined) return true;
    if (minCol !== undefined) {
      const col = letterToColIndex(change.cell.replace(/\d+/g, ''));
      if (col >= minCol) return false;
    }
    if (minRow !== undefined) {
      const row = Number.parseInt(change.cell.replace(/[A-Z]+/i, ''), 10);
      if (row >= minRow) return false;
    }
    return true;
  });
}

function cellRestoreActions(sheetName: string, cells: StructuralCellSnapshot[]): Action[] {
  const actions: Action[] = [];
  for (const cell of cells) {
    const format = buildRestoreFormat(cell);
    if (cell.formula && cell.formula.startsWith('=')) {
      actions.push({
        type: 'SET_FORMULA',
        sheetName,
        address: cell.address,
        formula: cell.formula,
        format,
        explicitOverwriteConfirmed: true,
      } as Action);
    } else if (cell.value !== null && cell.value !== undefined) {
      actions.push({
        type: 'SET_CELL',
        sheetName,
        address: cell.address,
        value: cell.value as string | number | boolean | null,
        format,
        explicitOverwriteConfirmed: true,
      } as Action);
    }
  }
  return actions;
}

/**
 * Build inverse actions for captured structural ops, split into `pre`/`post` relative to
 * the cell-level inverse actions from `beforeStateToInverseActions`:
 * - `pre` recreates what this change set removed (a deleted sheet, a deleted column — via
 *   ADD_SHEET/INSERT_COLUMN + its captured cell restore) *before* any cell-level inverse
 *   tries to write into it.
 * - `post` removes what this change set added (a sheet, a column) *after* cell-level
 *   inverses have run, so nothing tries to write into something already removed.
 * Ops are undone newest-first, matching standard undo-stack semantics. The real, dispatched
 * `INSERT_COLUMN`/`DELETE_COLUMN` action shapes (`beforeColumn`, `columns: string[]`) are
 * used here, not virtualApply's more permissive `col`/numeric forms — these inverse actions
 * are returned to the frontend and applied for real, not just simulated.
 */
export function structuralOpsToInverseActions(
  ops: StructuralOp[],
): { pre: Action[]; post: Action[] } {
  const pre: Action[] = [];
  const post: Action[] = [];

  for (const op of [...ops].reverse()) {
    if (op.opType === 'ADD_SHEET') {
      post.push({ type: 'DELETE_SHEET', sheetName: op.sheetName } as Action);
      continue;
    }

    if (op.opType === 'DELETE_SHEET') {
      pre.push({ type: 'ADD_SHEET', name: op.sheetName } as Action);
      const cells = (op.params.cells as StructuralCellSnapshot[] | undefined) ?? [];
      pre.push(...cellRestoreActions(op.sheetName, cells));
      continue;
    }

    if (op.opType === 'INSERT_COLUMN') {
      const index = op.params.index as number;
      const count = (op.params.count as number) ?? 1;
      const columns = Array.from({ length: count }, (_, i) => colIndexToLetter(index + i));
      post.push({ type: 'DELETE_COLUMN', sheetName: op.sheetName, columns } as Action);
      continue;
    }

    if (op.opType === 'DELETE_COLUMN') {
      const index = op.params.index as number;
      const count = (op.params.count as number) ?? 1;
      pre.push({
        type: 'INSERT_COLUMN',
        sheetName: op.sheetName,
        beforeColumn: colIndexToLetter(index),
        count,
      } as Action);
      const cells = (op.params.cells as StructuralCellSnapshot[] | undefined) ?? [];
      pre.push(...cellRestoreActions(op.sheetName, cells));
      continue;
    }

    if (op.opType === 'INSERT_ROW') {
      const index = op.params.index as number; // 1-based row the insert landed on
      const count = (op.params.count as number) ?? 1;
      const rows = Array.from({ length: count }, (_, i) => index + i);
      // SheetActionPayload's `rows` field is typed `unknown[][]` (shared with a table-data
      // use elsewhere) even though the real DELETE_ROW wire shape is `rows: number[]` —
      // virtualApply.ts's own DELETE_ROW handling casts through the same mismatch.
      post.push({ type: 'DELETE_ROW', sheetName: op.sheetName, rows } as unknown as Action);
      continue;
    }

    if (op.opType === 'DELETE_ROW') {
      const index = op.params.index as number; // 1-based row that was deleted
      const count = (op.params.count as number) ?? 1;
      pre.push({
        type: 'INSERT_ROW',
        sheetName: op.sheetName,
        row: index - 1, // 0-based "insert after" position, per InsertRowAction
        count,
      } as Action);
      const cells = (op.params.cells as StructuralCellSnapshot[] | undefined) ?? [];
      pre.push(...cellRestoreActions(op.sheetName, cells));
      continue;
    }

    if (op.opType === 'CREATE_TABLE') {
      // Unwrap only — cell values/formats are untouched by CREATE_TABLE itself
      // (see TASKS.md #16), so there's nothing else to restore here.
      const tableName = op.params.tableName as string;
      post.push({ type: 'DELETE_TABLE', sheetName: op.sheetName, tableName } as Action);
      continue;
    }

    if (op.opType === 'MERGE_CELLS') {
      // Must run *before* the generic cell-level inverse: MERGE_CELLS discards every
      // non-top-left cell's value (captured correctly by the generic diff already, per
      // TASKS.md #66/#41), but those cells can't be individually SET again until the
      // range is unmerged first (see TASKS.md #17).
      const range = op.params.range as string;
      pre.push({ type: 'UNMERGE_CELLS', sheetName: op.sheetName, range } as Action);
      continue;
    }

    if (op.opType === 'UNMERGE_CELLS') {
      // UNMERGE_CELLS has no cell-value effect to simulate (real Excel doesn't
      // resurrect the values a *prior* merge discarded) — just restore the merge
      // structure itself. Placed in `post` so it runs after any other pending
      // cell-level restores in this batch, rather than risk re-merging (and
      // silently discarding) a value another part of the revert still needs to write.
      const range = op.params.range as string;
      post.push({ type: 'MERGE_CELLS', sheetName: op.sheetName, range } as Action);
      continue;
    }

    if (op.opType === 'CONDITIONAL_FORMAT') {
      // TASKS.md #40 — ruleId is only present if ChangeSetService.markApplied()
      // successfully patched it in from the frontend's apply-time id read-back.
      // Fail loudly here rather than silently emit a DELETE_CONDITIONAL_FORMAT
      // with no id (which would either throw deep inside Office.js or, worse,
      // silently no-op) — same fail-closed posture as #19's revert self-verification.
      const ruleId = op.params.ruleId as string | undefined;
      if (!ruleId) {
        throw new Error(
          `Cannot build revert for CONDITIONAL_FORMAT on sheet "${op.sheetName}" (range "${op.params.range as string}") — no ruleId was ever reported back at apply time.`,
        );
      }
      post.push({ type: 'DELETE_CONDITIONAL_FORMAT', sheetName: op.sheetName, ruleId } as Action);
      continue;
    }

    if (op.opType === 'CREATE_CHART') {
      // TASKS.md #15 — chartId is only present if ChangeSetService.markApplied()
      // successfully patched it in from the frontend's apply-time id read-back.
      // Fail loudly here rather than silently emit a DELETE_CHART with no id —
      // same fail-closed posture as CONDITIONAL_FORMAT's identical #40 mechanism.
      const chartId = op.params.chartId as string | undefined;
      if (!chartId) {
        throw new Error(
          `Cannot build revert for CREATE_CHART on sheet "${op.sheetName}" (sourceRange "${op.params.sourceRange as string}") — no chartId was ever reported back at apply time.`,
        );
      }
      post.push({ type: 'DELETE_CHART', sheetName: op.sheetName, chartId } as Action);
    }
  }

  return { pre, post };
}

/**
 * Full, unoptimized comparison of two shadow workbooks — every cell key present in either
 * one, not just the ones a particular `virtualApply` call happened to mark `changedCells`.
 * Used by revert self-verification (TASKS.md #19): a cell an inverse action simply forgot
 * to touch would never appear in `changedCells`, so the normal `generateDiff` fast path
 * would silently miss it. Forcing the fallback full-key-collection path here is deliberate.
 */
export function diffShadowsFully(expected: ShadowWorkbook, actual: ShadowWorkbook): CellChange[] {
  const forcedFullCompare: ShadowWorkbook = { ...actual, changedCells: new Set() };
  return generateDiff(expected, forcedFullCompare);
}
