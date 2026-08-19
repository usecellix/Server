import { compareSortValues } from '../agents/utils/sort-value.util';
import {
  buildOutputRows,
  filterDataRows,
  applyFilterOperator,
  resolveFilterColumnIndex,
  findMatchingRowOffsets,
} from '../agents/utils/range-filter.util';
import { buildAggregateTable } from '../agents/utils/aggregate-table.util';
import { Logger } from '@nestjs/common';
import { Action } from '../agents/types/agent.types';
import {
  colIndexToLetter,
  deepCloneShadow,
  letterToColIndex,
  shadowSheetToContext,
} from './shadowWorkbook';
import { ShadowCell, ShadowSheet, ShadowWorkbook } from './shadowWorkbook.types';
import { parseA1Range } from '../agents/utils/range-merge.util';
import { stripSheetPrefix } from '../agents/utils/range-address.util';

const logger = new Logger('VirtualApply');

export function virtualApply(shadow: ShadowWorkbook, actions: Action[]): ShadowWorkbook {
  const wb = deepCloneShadow(shadow);

  for (const action of actions) {
    try {
      applyAction(wb, action);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.warn(`VirtualApply: failed to apply ${action.type}: ${message}`);
    }
  }

  return wb;
}

function applyAction(wb: ShadowWorkbook, action: Action): void {
  switch (action.type) {
    case 'ADD_ROW':
      virtualAddRow(wb, action);
      break;
    case 'SET_CELL':
      virtualSetCellLegacy(wb, action);
      break;
    case 'SET_FORMULA':
      virtualSetFormulaLegacy(wb, action);
      break;
    case 'BATCH_SET':
      virtualBatchSet(wb, action);
      break;
    case 'DELETE_ROW':
      virtualDeleteRowLegacy(wb, action);
      break;
    case 'INSERT_ROW':
      virtualInsertRowLegacy(wb, action);
      break;
    case 'INSERT_COLUMN':
      virtualInsertColumnLegacy(wb, action);
      break;
    case 'DELETE_COLUMN':
      virtualDeleteColumnLegacy(wb, action);
      break;
    case 'CREATE_SHEET':
    case 'ADD_SHEET':
      virtualAddSheet(wb, action.sheetName ?? action.name ?? 'Sheet', action.copyFrom);
      break;
    case 'RENAME_SHEET':
      virtualRenameSheet(
        wb,
        action.sheetName ?? action.oldName ?? '',
        action.newSheetName ?? action.newName ?? '',
      );
      break;
    case 'COPY_SHEET':
      virtualCopySheet(wb, action);
      break;
    case 'DELETE_SHEET':
      virtualDeleteSheet(wb, action.sheetName ?? '');
      break;
    case 'DEFINE_NAMED_RANGE':
      if (action.name && action.formula) {
        wb.namedRanges.set(action.name, action.formula);
      }
      break;
    case 'WRITE_TABLE':
      virtualWriteTable(wb, action);
      break;
    case 'SORT_RANGE':
      virtualSortRange(wb, action);
      break;
    case 'COPY_FILTERED_RANGE':
      virtualCopyFilteredRange(wb, action);
      break;
    case 'FORMAT_MATCHING_ROWS':
      // Format-only — shadow workbook has no fill state to update.
      break;
    case 'MOVE_RANGE':
      virtualMoveRange(wb, action);
      break;
    case 'AGGREGATE_TABLE':
      virtualAggregateTable(wb, action);
      break;
    case 'FILL_DOWN':
      virtualFillDown(wb, action);
      break;
    case 'FILL_RIGHT':
      virtualFillRight(wb, action);
      break;
    // ---- Deliberately not simulated — see virtual-apply-catalog.ts for the
    // documented reason behind every one of these, enforced by
    // test/virtual-apply-catalog.spec.ts so a new action type can't silently
    // land here without a recorded decision (TASKS.md #41). ----
    case 'FORMAT_RANGE':
    case 'AUTOFIT_COLUMNS':
    case 'CREATE_TABLE':
    case 'CREATE_CHART':
    case 'UPDATE_CHART':
    case 'CLARIFY':
    case 'CHECKPOINT':
    case 'HIGHLIGHT_CELL':
    case 'CLEAR_CELL':
    case 'HIDE_ROW':
    case 'UNHIDE_ROW':
    case 'SHOW_ROW':
    case 'HIDE_COLUMN':
    case 'UNHIDE_COLUMN':
    case 'SHOW_COLUMN':
    case 'SET_ROW_HEIGHT':
    case 'SET_COLUMN_WIDTH':
    case 'AUTO_FILTER':
    case 'FREEZE_PANES':
    case 'UNFREEZE_PANES':
    case 'HIDE_SHEET':
    case 'SHOW_SHEET':
    case 'SET_SHEET_COLOR':
    case 'PROTECT_SHEET':
    case 'UNPROTECT_SHEET':
    case 'SET_ZOOM':
    case 'ADD_COMMENT':
    case 'DELETE_COMMENT':
    case 'CLEAR_FORMAT':
    case 'UNMERGE_CELLS':
      break;
    // ---- TASKS.md #66: real simulations, found as gaps during #41's audit ----
    case 'CLEAR_CONTENT':
      virtualClearRange(wb, action, false);
      break;
    case 'CLEAR_ALL':
      virtualClearRange(wb, action, true);
      break;
    case 'SET_MATCHING_ROWS':
      virtualSetMatchingRows(wb, action);
      break;
    case 'MERGE_CELLS':
      virtualMergeCells(wb, action);
      break;
    default:
      break;
  }
}

function getSheet(wb: ShadowWorkbook, sheetName: string): ShadowSheet | undefined {
  return wb.sheets.get(sheetName);
}

function ensureSheet(wb: ShadowWorkbook, sheetName: string): ShadowSheet {
  let sheet = wb.sheets.get(sheetName);
  if (!sheet) {
    sheet = {
      name: sheetName,
      cells: new Map(),
      rowCount: 0,
      columnCount: 0,
      structure: 'unknown',
    };
    wb.sheets.set(sheetName, sheet);
  }
  return sheet;
}

function markChanged(wb: ShadowWorkbook, sheetName: string, address: string): void {
  wb.changedCells.add(`${sheetName}!${address}`);
}

function virtualSetCell(
  wb: ShadowWorkbook,
  sheetName: string,
  address: string,
  value: unknown,
  formula: string,
  numberFormat?: string,
): void {
  const sheet = ensureSheet(wb, sheetName);
  const existing = sheet.cells.get(address) ?? {
    value: null,
    formula: '',
    numberFormat: 'General',
  };
  sheet.cells.set(address, {
    ...existing,
    value: formula ? `[formula: ${formula}]` : value,
    formula,
    ...(numberFormat !== undefined ? { numberFormat } : {}),
  });
  markChanged(wb, sheetName, address);
  updateSheetBounds(sheet);
}

function virtualSetCellLegacy(wb: ShadowWorkbook, action: Action): void {
  if (action.address && action.sheetName) {
    virtualSetCell(
      wb,
      action.sheetName,
      action.address,
      action.value ?? null,
      '',
      action.format?.numberFormat,
    );
    return;
  }
  if (action.row === undefined || action.col === undefined) return;
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const address = `${colIndexToLetter(action.col)}${action.row + 1}`;
  virtualSetCell(wb, sheetName, address, action.value ?? null, '', action.format?.numberFormat);
}

function virtualSetFormulaLegacy(wb: ShadowWorkbook, action: Action): void {
  if (action.address && action.sheetName) {
    virtualSetCell(
      wb,
      action.sheetName,
      action.address,
      null,
      action.formula ?? '',
      action.format?.numberFormat,
    );
    return;
  }
  if (action.row === undefined || action.col === undefined || !action.formula) return;
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const address = `${colIndexToLetter(action.col)}${action.row + 1}`;
  virtualSetCell(wb, sheetName, address, null, action.formula, action.format?.numberFormat);
}

function virtualBatchSet(wb: ShadowWorkbook, action: Action): void {
  if (!action.sheetName || !Array.isArray(action.operations)) return;
  for (const op of action.operations) {
    virtualSetCell(
      wb,
      action.sheetName,
      op.address,
      op.value ?? null,
      op.formula ?? '',
      op.format?.numberFormat,
    );
  }
}

function virtualAddRow(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = ensureSheet(wb, sheetName);

  if (typeof action.afterRow === 'number' && Array.isArray(action.values)) {
    virtualAddRowAfter(wb, sheet, sheetName, action.afterRow, action.values);
    return;
  }

  if (!Array.isArray(action.data)) return;
  const insertAfter = sheet.rowCount > 0 ? sheet.rowCount : 0;
  virtualAddRowAfter(wb, sheet, sheetName, insertAfter, action.data);
}

function virtualAddRowAfter(
  wb: ShadowWorkbook,
  sheet: ShadowSheet,
  sheetName: string,
  afterRow: number,
  values: unknown[],
): void {
  const newCells = new Map<string, ShadowCell>();
  for (const [addr, cell] of sheet.cells) {
    const col = addr.replace(/\d+/g, '');
    const row = Number.parseInt(addr.replace(/[A-Z]+/i, ''), 10);
    if (row > afterRow) {
      newCells.set(`${col}${row + 1}`, cell);
    } else {
      newCells.set(addr, cell);
    }
  }

  values.forEach((val, i) => {
    const addr = `${colIndexToLetter(i)}${afterRow + 1}`;
    newCells.set(addr, { value: val, formula: '', numberFormat: 'General' });
    markChanged(wb, sheetName, addr);
  });

  sheet.cells = newCells;
  sheet.rowCount += 1;
  sheet.columnCount = Math.max(sheet.columnCount, values.length);
}

function virtualDeleteRowLegacy(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;

  if (Array.isArray(action.rowNumbers) && action.rowNumbers.length > 0) {
    virtualDeleteRows(wb, sheetName, action.rowNumbers);
    return;
  }

  if (Array.isArray(action.rows) && action.rows.every((r) => typeof r === 'number')) {
    virtualDeleteRows(wb, sheetName, action.rows as number[]);
    return;
  }

  if (action.row !== undefined) {
    virtualDeleteRows(wb, sheetName, [action.row + 1]);
  }
}

function virtualDeleteRows(wb: ShadowWorkbook, sheetName: string, rows: number[]): void {
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  const rowSet = new Set(rows);
  const newCells = new Map<string, ShadowCell>();
  let offset = 0;
  const maxRow = Math.max(sheet.rowCount, ...rows);

  for (let r = 1; r <= maxRow; r += 1) {
    if (rowSet.has(r)) {
      offset += 1;
      continue;
    }
    for (let c = 0; c < sheet.columnCount; c += 1) {
      const oldAddr = `${colIndexToLetter(c)}${r}`;
      const newAddr = `${colIndexToLetter(c)}${r - offset}`;
      const cell = sheet.cells.get(oldAddr);
      if (cell) newCells.set(newAddr, cell);
    }
  }

  sheet.cells = newCells;
  sheet.rowCount = Math.max(0, sheet.rowCount - rows.length);
}

function virtualInsertRowLegacy(wb: ShadowWorkbook, action: Action): void {
  if (action.row === undefined) return;
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const afterRow = action.position === 'below' ? action.row + 1 : action.row;
  const data = Array.isArray(action.data) ? action.data : [];
  const sheet = ensureSheet(wb, sheetName);
  virtualAddRowAfter(wb, sheet, sheetName, afterRow, data.length ? data : ['']);
}

function virtualInsertColumnLegacy(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  // Semantic INSERT_COLUMN: columnName + afterLastColumn / afterColumn
  const columnName =
    typeof action.columnName === 'string' ? action.columnName.trim() : '';
  const afterColumn =
    typeof action.afterColumn === 'string' ? action.afterColumn.trim() : '';
  const isSemantic =
    Boolean(columnName) &&
    (action.position === 'afterLastColumn' || Boolean(afterColumn));

  if (isSemantic) {
    let insertAt = sheet.columnCount;
    if (afterColumn) {
      const headerRow = 1; // Excel 1-based header in shadow keys
      let found = -1;
      for (let c = 0; c < sheet.columnCount; c += 1) {
        const addr = `${colIndexToLetter(c)}${headerRow}`;
        const cell = sheet.cells.get(addr);
        if (
          cell &&
          String(cell.value ?? '')
            .trim()
            .toLowerCase() === afterColumn.toLowerCase()
        ) {
          found = c;
          break;
        }
      }
      if (found >= 0) {
        insertAt = found + 1;
        virtualInsertColumn(wb, sheetName, insertAt, 1);
      }
    }

    const headerAddr = `${colIndexToLetter(insertAt)}1`;
    sheet.cells.set(headerAddr, {
      value: columnName,
      formula: '',
      numberFormat: 'General',
    });
    wb.changedCells.add(`${sheetName}!${headerAddr}`);

    if (typeof action.formula === 'string' && action.formula) {
      const dataRows = Math.max(sheet.rowCount - 1, 0);
      for (let r = 2; r <= dataRows + 1; r += 1) {
        const addr = `${colIndexToLetter(insertAt)}${r}`;
        const formula = action.formula.includes('{row}')
          ? action.formula.replace(/\{row\}/g, String(r))
          : action.formula.replace(/([A-Za-z]+)(\d+)/g, (_m, col: string, row: string) => {
              // Shift relative row refs from the template's base row to current row
              const base = Number(row);
              const delta = r - base;
              return `${col}${base + delta}`;
            });
        sheet.cells.set(addr, {
          value: null,
          formula,
          numberFormat: 'General',
        });
        wb.changedCells.add(`${sheetName}!${addr}`);
      }
    }

    sheet.columnCount = Math.max(sheet.columnCount, insertAt + 1);
    sheet.rowCount = Math.max(sheet.rowCount, 1);
    return;
  }

  const insertAt =
    action.beforeColumn !== undefined
      ? letterToColIndex(action.beforeColumn)
      : action.col ?? 0;
  const count = action.count ?? 1;
  virtualInsertColumn(wb, sheetName, insertAt, count);
}

function virtualInsertColumn(
  wb: ShadowWorkbook,
  sheetName: string,
  insertAt: number,
  count: number,
): void {
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  const newCells = new Map<string, ShadowCell>();
  for (const [addr, cell] of sheet.cells) {
    const col = letterToColIndex(addr.replace(/\d+/g, ''));
    const row = addr.replace(/[A-Z]+/i, '');
    if (col >= insertAt) {
      newCells.set(`${colIndexToLetter(col + count)}${row}`, cell);
    } else {
      newCells.set(addr, cell);
    }
  }

  sheet.cells = newCells;
  sheet.columnCount += count;
}

function virtualDeleteColumnLegacy(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  const rawColumns = Array.isArray(action.columns) ? action.columns : [];
  const resolvedIndices = new Set<number>();

  for (const raw of rawColumns) {
    const token = String(raw ?? '').trim();
    if (!token) continue;
    // Pure column letter(s) e.g. K, AA
    if (/^[A-Za-z]{1,3}$/.test(token)) {
      resolvedIndices.add(letterToColIndex(token));
      continue;
    }
    // Semantic header name e.g. "Payment Status"
    const headerIdx = findHeaderColumnIndex(sheet, token);
    if (headerIdx >= 0) {
      resolvedIndices.add(headerIdx);
    }
  }

  if (action.col !== undefined && Number.isFinite(action.col)) {
    resolvedIndices.add(action.col);
  }

  const cols = [...resolvedIndices].sort((a, b) => b - a);
  if (cols.length === 0) {
    logger.warn(
      `VirtualApply: DELETE_COLUMN had no resolvable columns (columns=${JSON.stringify(rawColumns)})`,
    );
    return;
  }

  for (const colIndex of cols) {
    const newCells = new Map<string, ShadowCell>();
    for (const [addr, cell] of sheet.cells) {
      const col = letterToColIndex(addr.replace(/\d+/g, ''));
      const row = addr.replace(/[A-Z]+/i, '');
      if (col > colIndex) {
        const newAddr = `${colIndexToLetter(col - 1)}${row}`;
        newCells.set(newAddr, cell);
        wb.changedCells.add(`${sheetName}!${newAddr}`);
      } else if (col < colIndex) {
        newCells.set(addr, cell);
      } else {
        // Deleted cell — mark old address so diff can see the removal
        wb.changedCells.add(`${sheetName}!${addr}`);
      }
    }
    sheet.cells = newCells;
    sheet.columnCount = Math.max(0, sheet.columnCount - 1);
  }
}

function findHeaderColumnIndex(sheet: ShadowSheet, headerName: string): number {
  const target = headerName.trim().toLowerCase();
  for (let c = 0; c < sheet.columnCount; c += 1) {
    const addr = `${colIndexToLetter(c)}1`;
    const cell = sheet.cells.get(addr);
    if (
      cell &&
      String(cell.value ?? '')
        .trim()
        .toLowerCase() === target
    ) {
      return c;
    }
  }
  return -1;
}

function virtualAddSheet(wb: ShadowWorkbook, name: string, copyFrom?: string): void {
  if (copyFrom && wb.sheets.has(copyFrom)) {
    const source = wb.sheets.get(copyFrom)!;
    wb.sheets.set(name, {
      name,
      cells: new Map(source.cells),
      rowCount: source.rowCount,
      columnCount: source.columnCount,
      structure: source.structure,
    });
  } else {
    wb.sheets.set(name, {
      name,
      cells: new Map(),
      rowCount: 0,
      columnCount: 0,
      structure: 'unknown',
    });
  }
}

function virtualRenameSheet(wb: ShadowWorkbook, oldName: string, newName: string): void {
  const sheet = wb.sheets.get(oldName);
  if (!sheet || !newName) return;
  sheet.name = newName;
  wb.sheets.set(newName, sheet);
  wb.sheets.delete(oldName);
  if (wb.activeSheetName === oldName) wb.activeSheetName = newName;
}

function virtualCopySheet(wb: ShadowWorkbook, action: Action): void {
  const sourceName = action.sourceName ?? action.sheetName;
  const newName = action.newSheetName ?? action.newName;
  if (!sourceName || !newName) return;
  virtualAddSheet(wb, newName, sourceName);
}

function virtualDeleteSheet(wb: ShadowWorkbook, sheetName: string): void {
  if (!sheetName || !wb.sheets.has(sheetName)) return;
  wb.sheets.delete(sheetName);
  if (wb.activeSheetName === sheetName) {
    const next = wb.sheets.keys().next();
    if (!next.done) wb.activeSheetName = next.value;
  }
}

/**
 * Shift relative A1 references by rowDelta/colDelta, leaving $-anchored
 * refs fixed — the same shape of logic virtualInsertColumnLegacy already
 * uses for its own formula template, generalized for fill-down/fill-right.
 * Naive regex, not a full formula parser — matches the sophistication
 * already accepted elsewhere in this file.
 */
function shiftFormulaRefs(formula: string, rowDelta: number, colDelta: number): string {
  return formula.replace(
    /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g,
    (_match, colDollar: string, col: string, rowDollar: string, row: string) => {
      const newCol =
        colDollar === '$' ? col : colIndexToLetter(letterToColIndex(col) + colDelta);
      const newRow = rowDollar === '$' ? row : String(Number(row) + rowDelta);
      return `${colDollar}${newCol}${rowDollar}${newRow}`;
    },
  );
}

/**
 * TASKS.md #42: populate the shadow with the correctly-shifted formula/value
 * for each filled cell, so FormulaValidatorService.checkPostApply's
 * error-string scan and reference-bounds re-check (both of which read
 * straight off the shadow, not the real workbook) can catch a bad reference
 * shift before the real Excel write — not by evaluating formulas (the shadow
 * never does that), just by making sure it actually holds what a fill would
 * really produce, instead of silently no-op'ing like before.
 */
function virtualFillDown(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  // Two distinct shapes reach virtualApply — this runs on the raw,
  // pre-normalization Action, before the wire ever sees it. The Executor's
  // own prompt examples use row/col/endRow (single column, e.g.
  // {"type":"FILL_DOWN","col":3,"row":1,"endRow":847}); sourceRange/
  // targetRange (A1 strings, possibly multi-column) is the canonical
  // shared/action.types.ts shape used elsewhere. Handling only one of these
  // silently no-ops for whichever shape wasn't checked — confirmed the
  // row/col/endRow shape is what the Executor is actually told to emit, so
  // it must not be the unhandled one.
  let srcRow: number;
  let srcColStart: number;
  let targetColStart: number;
  let targetColEnd: number;
  let targetRowStart: number;
  let targetRowEnd: number;

  if (action.sourceRange && action.targetRange) {
    const sourceBounds = parseA1Range(stripSheetPrefix(action.sourceRange));
    const targetBounds = parseA1Range(stripSheetPrefix(action.targetRange));
    if (!sourceBounds || !targetBounds) return;
    srcRow = sourceBounds.startRow;
    srcColStart = sourceBounds.startCol;
    targetColStart = targetBounds.startCol;
    targetColEnd = targetBounds.endCol;
    targetRowStart = targetBounds.startRow;
    targetRowEnd = targetBounds.endRow;
  } else if (
    typeof action.row === 'number' &&
    typeof action.col === 'number' &&
    typeof action.endRow === 'number'
  ) {
    srcRow = action.row;
    srcColStart = action.col;
    targetColStart = action.col;
    targetColEnd = action.col;
    targetRowStart = action.row + 1;
    targetRowEnd = action.endRow;
  } else {
    return;
  }

  for (let r = targetRowStart; r <= targetRowEnd; r += 1) {
    for (let c = targetColStart; c <= targetColEnd; c += 1) {
      const srcCol = srcColStart + (c - targetColStart);
      const srcCell = sheet.cells.get(`${colIndexToLetter(srcCol)}${srcRow + 1}`);
      if (!srcCell) continue;

      const rowDelta = r - srcRow;
      const isFormula = Boolean(srcCell.formula && srcCell.formula.startsWith('='));
      const formula = isFormula ? shiftFormulaRefs(srcCell.formula, rowDelta, 0) : '';
      const address = `${colIndexToLetter(c)}${r + 1}`;

      sheet.cells.set(address, {
        value: isFormula ? null : srcCell.value,
        formula,
        numberFormat: srcCell.numberFormat,
      });
      markChanged(wb, sheetName, address);
    }
  }
  updateSheetBounds(sheet);
}

/**
 * Resolve a rectangular span from either the row/col/rowCount/colCount shape
 * (confirmed via cellix-system-prompt.ts's own examples — the same
 * pre-normalization shape discovered for FILL_DOWN — to be what
 * MERGE_CELLS/CLEAR_CONTENT/CLEAR_ALL actually carry) or an A1 range
 * string, whichever the action has.
 */
function resolveSpan(
  action: Action,
): { startRow: number; startCol: number; endRow: number; endCol: number } | null {
  if (typeof action.row === 'number' && typeof action.col === 'number') {
    const rowCount =
      typeof action.rowCount === 'number' && action.rowCount > 0 ? action.rowCount : 1;
    const colCount =
      typeof action.colCount === 'number' && action.colCount > 0 ? action.colCount : 1;
    return {
      startRow: action.row,
      startCol: action.col,
      endRow: action.row + rowCount - 1,
      endCol: action.col + colCount - 1,
    };
  }
  if (action.range) {
    const bounds = parseA1Range(stripSheetPrefix(action.range));
    if (bounds) return bounds;
  }
  return null;
}

/** TASKS.md #66 — CLEAR_CONTENT / CLEAR_ALL: actually clear values (and format, for CLEAR_ALL) in the shadow. */
function virtualClearRange(wb: ShadowWorkbook, action: Action, clearFormat: boolean): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  const span = resolveSpan(action);
  if (!span) return;

  for (let r = span.startRow; r <= span.endRow; r += 1) {
    for (let c = span.startCol; c <= span.endCol; c += 1) {
      const address = `${colIndexToLetter(c)}${r + 1}`;
      const existing = sheet.cells.get(address);
      if (!existing) continue; // nothing there to clear
      sheet.cells.set(address, {
        value: null,
        formula: '',
        numberFormat: clearFormat ? 'General' : existing.numberFormat,
      });
      markChanged(wb, sheetName, address);
    }
  }
}

/** TASKS.md #66 — SET_MATCHING_ROWS: write `value` into `targetColumn` for every row matching `filter` (or all data rows if omitted). */
function virtualSetMatchingRows(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = getSheet(wb, sheetName);
  if (!sheet || !action.targetColumn) return;

  const rangeStr = action.range ?? action.sourceRange;
  if (!rangeStr) return;
  const bounds = parseA1Range(stripSheetPrefix(rangeStr));
  if (!bounds) return;

  const rows = readRangeValues(sheet, rangeStr);
  if (rows.length === 0) return;

  const hasHeaders = action.hasHeaders !== false;
  const headerRow = hasHeaders ? rows[0] : null;
  if (!headerRow) return; // matching by header name requires a header row, same as the real Office.js handler

  let targetColIndex: number;
  try {
    targetColIndex = resolveFilterColumnIndex(headerRow, action.targetColumn);
  } catch {
    return;
  }

  const matchOffsets = action.filter
    ? findMatchingRowOffsets(rows, hasHeaders, action.filter)
    : rows.map((_, i) => i).filter((i) => i >= (hasHeaders ? 1 : 0));

  for (const offset of matchOffsets) {
    const address = `${colIndexToLetter(bounds.startCol + targetColIndex)}${bounds.startRow + offset + 1}`;
    virtualSetCell(wb, sheetName, address, action.value ?? null, '');
  }
}

/**
 * TASKS.md #66 — MERGE_CELLS: real Excel keeps only the top-left cell's
 * value and discards the rest. Modeling that (not "no change") is what lets
 * the existing generic before/after diff — the same mechanism the preview
 * UI already surfaces to the user — show the discarded values as a real
 * change instead of them silently vanishing with no signal at all.
 */
function virtualMergeCells(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  const span = resolveSpan(action);
  if (!span) return;
  if (span.startRow === span.endRow && span.startCol === span.endCol) return; // single cell, nothing to merge

  for (let r = span.startRow; r <= span.endRow; r += 1) {
    for (let c = span.startCol; c <= span.endCol; c += 1) {
      if (r === span.startRow && c === span.startCol) continue; // top-left survives untouched
      const address = `${colIndexToLetter(c)}${r + 1}`;
      const existing = sheet.cells.get(address);
      if (!existing) continue;
      sheet.cells.set(address, { value: null, formula: '', numberFormat: existing.numberFormat });
      markChanged(wb, sheetName, address);
    }
  }
}

/** TASKS.md #42 — same purpose as virtualFillDown, shifting columns instead of rows. */
function virtualFillRight(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  let srcRow: number;
  let srcCol: number;
  let endCol: number;

  if (action.sourceRange && action.targetRange) {
    const sourceBounds = parseA1Range(stripSheetPrefix(action.sourceRange));
    const targetBounds = parseA1Range(stripSheetPrefix(action.targetRange));
    if (!sourceBounds || !targetBounds) return;
    srcRow = sourceBounds.startRow;
    srcCol = sourceBounds.startCol;
    endCol = targetBounds.endCol;
  } else if (
    typeof action.row === 'number' &&
    typeof action.col === 'number' &&
    typeof action.endCol === 'number'
  ) {
    srcRow = action.row;
    srcCol = action.col;
    endCol = action.endCol;
  } else {
    return;
  }

  const srcCell = sheet.cells.get(`${colIndexToLetter(srcCol)}${srcRow + 1}`);
  if (!srcCell) return;

  for (let c = srcCol + 1; c <= endCol; c += 1) {
    const colDelta = c - srcCol;
    const isFormula = Boolean(srcCell.formula && srcCell.formula.startsWith('='));
    const formula = isFormula ? shiftFormulaRefs(srcCell.formula, 0, colDelta) : '';
    const address = `${colIndexToLetter(c)}${srcRow + 1}`;

    sheet.cells.set(address, {
      value: isFormula ? null : srcCell.value,
      formula,
      numberFormat: srcCell.numberFormat,
    });
    markChanged(wb, sheetName, address);
  }
  updateSheetBounds(sheet);
}

function virtualWriteTable(wb: ShadowWorkbook, action: Action): void {
  if (!Array.isArray(action.headers) || !Array.isArray(action.rows)) return;
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = ensureSheet(wb, sheetName);
  const allRows = [action.headers, ...action.rows];
  allRows.forEach((row, rowIdx) => {
    if (!Array.isArray(row)) return;
    row.forEach((val, colIdx) => {
      const addr = `${colIndexToLetter(colIdx)}${rowIdx + 1}`;
      virtualSetCell(wb, sheetName, addr, val, '');
    });
  });
  sheet.rowCount = Math.max(sheet.rowCount, allRows.length);
  sheet.columnCount = Math.max(
    sheet.columnCount,
    ...allRows.map((r) => (Array.isArray(r) ? r.length : 0)),
    0,
  );
}

function virtualSortRange(wb: ShadowWorkbook, action: Action): void {
  const sheetName = action.sheetName ?? wb.activeSheetName;
  const sheet = getSheet(wb, sheetName);
  if (!sheet) return;

  const snapshot = shadowSheetToContext(sheet);
  if (snapshot.rowCount < 2) return;

  const key = action.key ?? 0;
  const ascending = action.ascending ?? true;
  const hasHeaders = action.hasHeaders ?? true;
  const headerRow = hasHeaders ? snapshot.values[0] : null;
  const dataStart = hasHeaders ? 1 : 0;
  const dataRows = snapshot.values.slice(dataStart).map((row) => [...row]);

  // Compressed workbook context pads missing rows with nulls. Sorting that matrix
  // produces garbage diffs (clears + wrong moves) while the client SORT_RANGE on the
  // live sheet is what actually matters — skip virtual mutation when data is sparse.
  const populatedDataRows = dataRows.filter((row) =>
    row.some((cell) => cell !== null && cell !== undefined && cell !== ''),
  );
  if (populatedDataRows.length < dataRows.length) {
    return;
  }

  const formulaRows = snapshot.formulas.slice(dataStart).map((row) => [...row]);
  const formatRows = snapshot.numberFormats.slice(dataStart).map((row) => [...row]);

  const indices = dataRows.map((_, index) => index);
  indices.sort((ia, ib) => {
    const cmp = compareSortValues(dataRows[ia]?.[key], dataRows[ib]?.[key]);
    return ascending ? cmp : -cmp;
  });

  const sortedValues = indices.map((i) => dataRows[i]);
  const sortedFormulas = indices.map((i) => formulaRows[i] ?? []);
  const sortedFormats = indices.map((i) => formatRows[i] ?? []);

  sheet.cells.clear();
  const allValueRows = headerRow ? [headerRow, ...sortedValues] : sortedValues;
  const allFormulaRows = hasHeaders
    ? [snapshot.formulas[0] ?? [], ...sortedFormulas]
    : sortedFormulas;
  const allFormatRows = hasHeaders
    ? [snapshot.numberFormats[0] ?? [], ...sortedFormats]
    : sortedFormats;

  for (let r = 0; r < allValueRows.length; r += 1) {
    const row = allValueRows[r] ?? [];
    for (let c = 0; c < row.length; c += 1) {
      const address = `${colIndexToLetter(c)}${r + 1}`;
      const formula = allFormulaRows[r]?.[c] ?? '';
      const value = row[c];
      sheet.cells.set(address, {
        value: formula && formula.startsWith('=') ? null : value,
        formula: formula && formula.startsWith('=') ? formula : '',
        numberFormat: allFormatRows[r]?.[c] ?? 'General',
      });
      markChanged(wb, sheetName, address);
    }
  }

  updateSheetBounds(sheet);
}

function updateSheetBounds(sheet: ShadowSheet): void {
  const addresses = Array.from(sheet.cells.keys());
  if (!addresses.length) return;
  sheet.rowCount = Math.max(
    sheet.rowCount,
    ...addresses.map((a) => Number.parseInt(a.replace(/[A-Z]+/i, ''), 10)),
  );
  sheet.columnCount = Math.max(
    sheet.columnCount,
    ...addresses.map((a) => letterToColIndex(a.replace(/\d+/g, '')) + 1),
  );
}

function parseDestStartCell(destStartCell: string): { row: number; col: number } | null {
  const parsed = parseA1Range(stripSheetPrefix(destStartCell));
  if (!parsed) return null;
  return { row: parsed.startRow, col: parsed.startCol };
}

function readRangeValues(
  sheet: ShadowSheet,
  sourceRange: string,
): unknown[][] {
  const bounds = parseA1Range(stripSheetPrefix(sourceRange));
  if (!bounds) {
    // Fall back to full sheet snapshot when range is invalid/empty
    return shadowSheetToContext(sheet).values;
  }

  const rows: unknown[][] = [];
  for (let r = bounds.startRow; r <= bounds.endRow; r += 1) {
    const row: unknown[] = [];
    for (let c = bounds.startCol; c <= bounds.endCol; c += 1) {
      const address = `${colIndexToLetter(c)}${r + 1}`;
      const cell = sheet.cells.get(address);
      row.push(cell?.value ?? null);
    }
    rows.push(row);
  }
  return rows;
}

function writeRowsAt(
  wb: ShadowWorkbook,
  sheetName: string,
  startRow: number,
  startCol: number,
  outputRows: unknown[][],
): void {
  const sheet = ensureSheet(wb, sheetName);
  for (let r = 0; r < outputRows.length; r += 1) {
    const row = outputRows[r] ?? [];
    for (let c = 0; c < row.length; c += 1) {
      const address = `${colIndexToLetter(startCol + c)}${startRow + r + 1}`;
      virtualSetCell(wb, sheetName, address, row[c] ?? null, '');
    }
  }
  sheet.rowCount = Math.max(sheet.rowCount, startRow + outputRows.length);
  sheet.columnCount = Math.max(
    sheet.columnCount,
    startCol + (outputRows.length ? Math.max(...outputRows.map((row) => row.length)) : 0),
  );
}

function virtualCopyFilteredRange(wb: ShadowWorkbook, action: Action): void {
  const sourceSheetName = action.sourceSheet ?? action.sheetName ?? wb.activeSheetName;
  const destSheetName = action.destSheet;
  const sourceRange = action.sourceRange ?? action.range;
  const destStartCell = action.destStartCell ?? 'A1';
  if (!destSheetName || !sourceRange) return;

  const sourceSheet = getSheet(wb, sourceSheetName);
  if (!sourceSheet) return;

  const rows = readRangeValues(sourceSheet, sourceRange);
  const hasHeaders = action.hasHeaders ?? true;
  const { headerRow, filteredRows } = filterDataRows(rows, hasHeaders, action.filter);
  const outputRows = buildOutputRows(headerRow, filteredRows);
  if (outputRows.length === 0) return;

  const dest = parseDestStartCell(destStartCell);
  if (!dest) return;

  writeRowsAt(wb, destSheetName, dest.row, dest.col, outputRows);

  if (action.mode === 'move' && action.filter && hasHeaders && headerRow) {
    clearMatchedSourceRows(wb, sourceSheetName, sourceRange, hasHeaders, action.filter);
  }
}

function virtualMoveRange(wb: ShadowWorkbook, action: Action): void {
  const sourceSheetName = action.sourceSheet ?? action.sheetName ?? wb.activeSheetName;
  const destSheetName = action.destSheet;
  const sourceRange = action.sourceRange ?? action.range;
  const destStartCell = action.destStartCell ?? 'A1';
  if (!destSheetName || !sourceRange) return;

  const sourceSheet = getSheet(wb, sourceSheetName);
  if (!sourceSheet) return;

  const rows = readRangeValues(sourceSheet, sourceRange);
  if (rows.length === 0) return;

  const dest = parseDestStartCell(destStartCell);
  if (!dest) return;

  writeRowsAt(wb, destSheetName, dest.row, dest.col, rows);

  const bounds = parseA1Range(stripSheetPrefix(sourceRange));
  if (!bounds) return;
  for (let r = bounds.startRow; r <= bounds.endRow; r += 1) {
    for (let c = bounds.startCol; c <= bounds.endCol; c += 1) {
      const address = `${colIndexToLetter(c)}${r + 1}`;
      virtualSetCell(wb, sourceSheetName, address, null, '');
    }
  }
}

function clearMatchedSourceRows(
  wb: ShadowWorkbook,
  sourceSheetName: string,
  sourceRange: string,
  hasHeaders: boolean,
  filter: NonNullable<Action['filter']>,
): void {
  const sourceSheet = getSheet(wb, sourceSheetName);
  if (!sourceSheet) return;

  const bounds = parseA1Range(stripSheetPrefix(sourceRange));
  if (!bounds) return;

  const rows = readRangeValues(sourceSheet, sourceRange);
  const headerRow = hasHeaders && rows.length > 0 ? rows[0] : null;
  if (!headerRow) return;

  const colIndex = headerRow.findIndex(
    (cell) => String(cell ?? '').trim().toLowerCase() === filter.column.trim().toLowerCase(),
  );
  if (colIndex === -1) return;

  const dataStart = hasHeaders ? 1 : 0;
  // Clear matched data rows bottom-to-top
  for (let i = rows.length - 1; i >= dataStart; i -= 1) {
    const row = rows[i];
    if (!row) continue;
    if (!applyFilterOperator(row[colIndex], filter)) continue;
    const absoluteRow = bounds.startRow + i;
    for (let c = bounds.startCol; c <= bounds.endCol; c += 1) {
      const address = `${colIndexToLetter(c)}${absoluteRow + 1}`;
      virtualSetCell(wb, sourceSheetName, address, null, '');
    }
  }
}

function virtualAggregateTable(wb: ShadowWorkbook, action: Action): void {
  const sourceSheetName = action.sourceSheet ?? action.sheetName ?? wb.activeSheetName;
  const destSheetName = action.destSheet;
  const sourceRange = action.sourceRange ?? action.range;
  const destStartCell = action.destStartCell ?? 'A1';
  if (!destSheetName || !sourceRange || !action.groupByColumn || !action.aggregations?.length) {
    return;
  }

  const sourceSheet = getSheet(wb, sourceSheetName);
  if (!sourceSheet) return;

  const rows = readRangeValues(sourceSheet, sourceRange);
  const outputRows = buildAggregateTable({
    rows,
    hasHeaders: action.hasHeaders !== false,
    groupByColumn: action.groupByColumn,
    groupByTransform: action.groupByTransform,
    aggregations: action.aggregations,
    sortBy: action.sortBy,
    topN: action.topN,
  });
  if (outputRows.length === 0) return;

  const dest = parseDestStartCell(destStartCell);
  if (!dest) return;
  writeRowsAt(wb, destSheetName, dest.row, dest.col, outputRows);
}

