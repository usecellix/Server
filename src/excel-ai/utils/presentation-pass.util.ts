import { SheetActionPayload } from '../types/sheet-actions.types';
import { WorkbookContext } from '../../types/cellix.types';
import { sheetsCreatedInBatch } from './sheet-header-state.util';

/**
 * Deterministic presentation pass — TASKS.md #138.
 *
 * The planner reliably produces the right *content* for a workbook build (12
 * month sheets, a Main summary, cross-sheet SUM/SUMIF formulas) and reliably
 * produces no *presentation* at all: the reported 189-action monthly-ledger
 * build was 188 SET_CELL and one CREATE_CHART, with zero FORMAT_RANGE,
 * AUTOFIT_COLUMNS or FREEZE_PANES — despite all three being supported action
 * types. The result is correct and looks unfinished: default column widths
 * clipping "Rate Per Ni", raw `0` where a currency belongs, no header styling.
 *
 * Styling is applied in code rather than asked of the model on purpose. The
 * quality users compare us against comes from being *consistent* across 13
 * sheets, which is exactly what an LLM re-improvising a house style 13 times
 * will not give. It also costs no tokens and cannot truncate.
 *
 * Two hard constraints:
 *
 *  1. **Layout-preserving.** This pass never moves a cell, inserts a row, or
 *     changes an anchor. The planner computes exact anchors (the consolidated
 *     header at `Main!A18` is derived from the Monthly Totals table ending at
 *     row 16); shifting anything would silently break that arithmetic.
 *  2. **Created sheets only.** Styling is applied solely to sheets this batch
 *     creates. An existing sheet has the user's own formatting, and
 *     `planner.prompt.ts` is explicit that unrequested formatting changes to
 *     existing cells are out of bounds — including inventing a number format.
 */

/** Minimum labelled cells on one row before it reads as a table header. */
const MIN_HEADER_CELLS = 3;

/** Rows of number formatting laid down below a header, for data not yet entered. */
const FORMAT_RUNWAY_ROWS = 200;

const HEADER_FILL = '#2F5597';
const HEADER_FONT = '#FFFFFF';
const TITLE_FONT = '#2F5597';

/**
 * Column headers whose values are money.
 *
 * The first version of this list was written against one hospitality prompt
 * (amount / paid / pending / rate / total) and silently did nothing for a
 * payroll, invoicing, expense or tax workbook — "Salary", "Commission", "GST",
 * "Freight", "Deposit" all fell through to General. Number formatting here is
 * additive (a miss means plain, never wrong), so the list is deliberately broad
 * across the business domains this product targets. TASKS.md #158.
 */
const CURRENCY_HEADER_WORDS = [
  // generic money
  'amount', 'total', 'subtotal', 'sum', 'value', 'price', 'cost', 'rate',
  'balance', 'due', 'paid', 'unpaid', 'pending', 'outstanding', 'received',
  'payable', 'receivable', 'credit', 'debit', 'net', 'gross',
  // revenue / expense
  'revenue', 'income', 'earnings', 'turnover', 'sales', 'expense', 'expenses',
  'spend', 'budget', 'margin', 'profit', 'loss', 'fee', 'fees', 'charge', 'charges',
  // payroll
  'salary', 'wage', 'wages', 'pay', 'payment', 'bonus', 'commission',
  'allowance', 'deduction', 'reimbursement', 'stipend',
  // billing / tax, including the India-specific terms this product targets
  'invoice', 'bill', 'billed', 'discount', 'refund', 'deposit', 'advance',
  'tax', 'gst', 'cgst', 'sgst', 'igst', 'vat', 'tds', 'cess', 'duty',
  'premium', 'freight', 'shipping', 'insurance', 'interest', 'principal',
  'emi', 'instalment', 'installment',
];

const CURRENCY_HEADER_PATTERN = new RegExp(
  `\\b(${CURRENCY_HEADER_WORDS.join('|')})\\b`,
  'i',
);

/** Column headers whose values are whole-number counts, not money. */
const COUNT_HEADER_WORDS = [
  'qty', 'quantity', 'count', 'number', 'nos', 'unit', 'units', 'item', 'items',
  'pcs', 'pieces', 'night', 'nights', 'day', 'days', 'hour', 'hours',
  'week', 'weeks', 'guest', 'guests', 'people', 'person', 'persons',
  'headcount', 'employee', 'employees', 'student', 'students',
  'room', 'rooms', 'seat', 'seats', 'booking', 'bookings', 'order', 'orders',
  'ticket', 'tickets', 'attendance', 'stock', 'inventory',
];

const COUNT_HEADER_PATTERN = new RegExp(
  `\\b(${COUNT_HEADER_WORDS.join('|')})\\b`,
  'i',
);

/**
 * Headers that must NEVER receive a numeric format, checked before both lists
 * above and winning outright.
 *
 * Two distinct hazards, both found while broadening the word lists:
 *
 *  - **Text columns that contain a money/count word.** "Guest Name" contains
 *    the whole word "Guest"; "Payment Status" contains "Payment". Applying
 *    `#,##0` to a name column is visibly wrong.
 *  - **Identifiers that look numeric but are not quantities.** "Unit No",
 *    "Invoice Number", "Account Code", "GST No" — thousands-separating unit
 *    1001 into "1,001" corrupts the meaning of the value.
 *
 * Conservative by design: a false veto costs plain formatting, which is exactly
 * the old behaviour and never wrong. A false positive corrupts how a real value
 * reads. TASKS.md #158.
 */
const NON_NUMERIC_HEADER_WORDS = [
  'name', 'names', 'description', 'desc', 'remark', 'remarks', 'note', 'notes',
  'comment', 'comments', 'status', 'source', 'type', 'category', 'address',
  'email', 'phone', 'mobile', 'contact', 'code', 'ref', 'reference',
  'id', 'no', 'number', 'account', 'gstin', 'pan', 'method', 'mode',
  'currency', 'label', 'title', 'notes',
];

const NON_NUMERIC_HEADER_PATTERN = new RegExp(
  `\\b(${NON_NUMERIC_HEADER_WORDS.join('|')})\\b`,
  'i',
);

/** Currency codes we will honour when the user or workbook names one. */
const CURRENCY_FORMATS: Record<string, string> = {
  INR: '₹#,##0',
  USD: '$#,##0.00',
  EUR: '€#,##0.00',
  GBP: '£#,##0.00',
};

const CURRENCY_MENTIONS: [RegExp, string][] = [
  [/₹|\brs\.?\b|\binr\b|\brupees?\b/i, 'INR'],
  [/\$|\busd\b|\bdollars?\b/i, 'USD'],
  [/€|\beur\b|\beuros?\b/i, 'EUR'],
  [/£|\bgbp\b|\bpounds?\b/i, 'GBP'],
];

/** Neutral fallback — a thousands separator, no invented currency symbol. */
const NEUTRAL_NUMBER_FORMAT = '#,##0.00';

export interface PresentationOptions {
  /** The user's own words — the only place a currency may be inferred from. */
  userMessage?: string;
  context?: WorkbookContext;
}

export interface HeaderRun {
  sheetName: string;
  row: number;
  startCol: number;
  endCol: number;
  labels: Map<number, string>;
}

/**
 * Resolve the number format for money columns.
 *
 * Never guessed. A currency is used only when the user named one, or when the
 * workbook already uses one — otherwise a neutral thousands format, which
 * improves readability without asserting a currency the user never mentioned.
 */
export function resolveCurrencyFormat(options: PresentationOptions): string {
  const message = options.userMessage ?? '';
  for (const [pattern, code] of CURRENCY_MENTIONS) {
    if (pattern.test(message)) return CURRENCY_FORMATS[code];
  }

  for (const sheet of options.context?.sheets ?? []) {
    for (const meta of sheet.columnMeta ?? []) {
      const existing = String((meta as { numberFormat?: string }).numberFormat ?? '');
      for (const [pattern, code] of CURRENCY_MENTIONS) {
        if (pattern.test(existing)) return CURRENCY_FORMATS[code];
      }
    }
  }

  return NEUTRAL_NUMBER_FORMAT;
}

/** Parse an A1 address into 0-based row/col. Returns null for anything else. */
function parseA1Address(address: string): { row: number; col: number } | null {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(address.trim());
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

/**
 * True when every element of an ADD_ROW's data looks like a column label —
 * non-empty strings, none numeric — the same test `detectHeaderRuns` already
 * applies per cell for SET_CELL.
 */
function looksLikeHeaderRow(data: unknown[]): boolean {
  if (data.length < MIN_HEADER_CELLS) return false;
  return data.every(
    (cell) => typeof cell === 'string' && cell.trim() !== '' && Number.isNaN(Number(cell)),
  );
}

/**
 * Find every table-header row this batch writes.
 *
 * Two shapes reach here for the identical intent: a header written as ten
 * SET_CELLs, and a header written as one ADD_ROW carrying a `data`/`values`
 * array — both are legitimate ways for the executor to express "write this
 * header row" (`convertHeaderRowWritesToAddRow` upstream in this same service
 * produces the second shape whenever the sheet already had a header to
 * protect, and the executor sometimes reaches for it directly). Recognizing
 * only the first meant a batch that mixed shapes across sheets — eleven months
 * as SET_CELL, one as ADD_ROW, observed live — styled and consolidated eleven
 * sheets and silently skipped the twelfth. See TASKS.md #143.
 *
 * ADD_ROW carries no row index (Office.js resolves it from the live used
 * range at apply time), so it is only trusted as a header when the target
 * sheet is one THIS batch creates — on a sheet that did not exist a moment
 * ago, the append is guaranteed to land at row 0.
 */
export function detectHeaderRuns(
  actions: SheetActionPayload[],
  createdSheets: Set<string>,
): HeaderRun[] {
  const byRow = new Map<string, HeaderRun>();

  const upsert = (sheetName: string, row: number, col: number, label: string) => {
    const key = `${sheetName.toLowerCase()}#${row}`;
    const existing = byRow.get(key);
    if (existing) {
      existing.startCol = Math.min(existing.startCol, col);
      existing.endCol = Math.max(existing.endCol, col);
      existing.labels.set(col, label);
    } else {
      byRow.set(key, { sheetName, row, startCol: col, endCol: col, labels: new Map([[col, label]]) });
    }
  };

  for (const action of actions) {
    const sheetName = String(action.sheetName ?? '').trim();
    if (!sheetName || !createdSheets.has(sheetName.toLowerCase())) continue;

    if (action.type === 'SET_CELL') {
      if (typeof action.row !== 'number' || typeof action.col !== 'number') continue;
      if (typeof action.value !== 'string' || action.value.trim() === '') continue;
      // A numeric-looking label is data, not a header.
      if (!Number.isNaN(Number(action.value))) continue;
      upsert(sheetName, action.row, action.col, action.value);
      continue;
    }

    if (action.type === 'ADD_ROW') {
      const data = (action.data ?? action.values) as unknown[] | undefined;
      if (!Array.isArray(data) || !looksLikeHeaderRow(data)) continue;
      data.forEach((cell, col) => upsert(sheetName, 0, col, cell as string));
      continue;
    }

    // Third shape: a header written as one BATCH_SET. The Executor picks
    // between SET_CELL, ADD_ROW and BATCH_SET run to run for the identical
    // intent, and a detector that understands only some of them silently skips
    // whole sheets — #143 was this same lesson for ADD_ROW. TASKS.md #157.
    if (action.type === 'BATCH_SET' && Array.isArray(action.operations)) {
      for (const op of action.operations) {
        if (typeof op?.address !== 'string') continue;
        const parsed = parseA1Address(op.address);
        if (!parsed) continue;
        if (typeof op.value !== 'string' || op.value.trim() === '') continue;
        if (!Number.isNaN(Number(op.value))) continue;
        upsert(sheetName, parsed.row, parsed.col, op.value);
      }
    }
  }

  return [...byRow.values()]
    .filter(
      (run) =>
        run.labels.size >= MIN_HEADER_CELLS &&
        // Contiguity, not just count. A live run put a KPI label row
        // (A1="Total Amount", C1="Paid Amount", E1="Pending Amount", with B1/D1
        // empty — their SUM formulas live on the row below) through this check:
        // 3 labels met MIN_HEADER_CELLS, and the label-count test alone doesn't
        // notice the span it inferred (A:E, 5 columns) is only 60% labelled. A
        // genuine table header has no gaps. The consequence was concrete: a
        // header band painted across the two empty cells, number formats landed
        // on the label columns instead of the value columns beside them, and —
        // worse — this became Main's numerically-topmost "header", so the two
        // REAL table headers on the same sheet lost their freeze/autofit
        // entirely (only one is chosen per sheet). See TASKS.md #144.
        run.labels.size === run.endCol - run.startCol + 1,
    )
    .sort((a, b) => a.sheetName.localeCompare(b.sheetName) || a.row - b.row);
}

/**
 * Append presentation actions for everything this batch builds.
 *
 * Returns the original actions followed by the styling ones — order matters:
 * `RichActionEngine` applies in array order, and autofit must run after the
 * values it measures exist.
 */
export function applyPresentationPass(
  actions: SheetActionPayload[],
  options: PresentationOptions = {},
): SheetActionPayload[] {
  const createdSheets = sheetsCreatedInBatch(actions);
  if (createdSheets.size === 0) return actions;

  const runs = detectHeaderRuns(actions, createdSheets);
  if (runs.length === 0) return actions;

  const currencyFormat = resolveCurrencyFormat(options);
  const styling: SheetActionPayload[] = [];
  const topHeaderRowBySheet = new Map<string, HeaderRun>();

  for (const run of runs) {
    const key = run.sheetName.toLowerCase();
    const current = topHeaderRowBySheet.get(key);
    if (!current || run.row < current.row) topHeaderRowBySheet.set(key, run);

    const colCount = run.endCol - run.startCol + 1;

    // Header band.
    styling.push({
      type: 'FORMAT_RANGE',
      sheetName: run.sheetName,
      row: run.row,
      col: run.startCol,
      rowCount: 1,
      colCount,
      format: {
        bold: true,
        fontColor: HEADER_FONT,
        fillColor: HEADER_FILL,
        horizontalAlignment: 'center',
        verticalAlignment: 'middle',
        wrapText: true,
        borders: 'all',
      },
    });

    styling.push({
      type: 'SET_ROW_HEIGHT',
      sheetName: run.sheetName,
      row: run.row,
      height: 28,
    });

    // Number formats for the columns under this header.
    //
    // The runway must stop before the NEXT table on the same sheet. Main stacks
    // a Monthly Totals table at row 4 and a consolidated transactions table at
    // row 18; a flat 200-row runway from the first one painted currency down
    // B5:D204, straight through the second table's Unit No / Guest / Guest Name
    // columns. Two conflicting number formats then landed on the same cells and
    // the later one won -- silently, since a number format never errors.
    // TASKS.md #164.
    const runwayRows = runwayFor(run, runs);

    for (const [col, label] of run.labels) {
      const numberFormat = numberFormatForLabel(label, currencyFormat);
      if (!numberFormat) continue;

      styling.push({
        type: 'FORMAT_RANGE',
        sheetName: run.sheetName,
        row: run.row + 1,
        col,
        rowCount: runwayRows,
        colCount: 1,
        format: { numberFormat, horizontalAlignment: 'right' },
      });
    }
  }

  // Per-sheet finishing: KPI band, title styling, frozen header, autofit.
  for (const [, run] of topHeaderRowBySheet) {
    styling.push(...styleKpiBand(actions, run, currencyFormat));

    const titleRow = findTitleRow(actions, run);
    if (titleRow !== null) {
      styling.push({
        type: 'FORMAT_RANGE',
        sheetName: run.sheetName,
        row: titleRow,
        col: 0,
        rowCount: 1,
        colCount: Math.max(run.endCol + 1, 1),
        format: { bold: true, fontSize: 14, fontColor: TITLE_FONT },
      });
    }

    styling.push({
      type: 'FREEZE_PANES',
      sheetName: run.sheetName,
      freezeRows: run.row + 1,
      freezeColumns: 0,
    });

    styling.push({ type: 'AUTOFIT_COLUMNS', sheetName: run.sheetName });
  }

  return [...actions, ...styling];
}

/**
 * A lone label written above the topmost header row reads as a sheet title
 * (e.g. Main's "Dashboard"). Only a single-cell row qualifies — a multi-cell
 * row above a header is another table, not a title.
 */
function findTitleRow(actions: SheetActionPayload[], header: HeaderRun): number | null {
  const rows = new Map<number, number>();

  // Counts LABEL cells only: a title row is a row above the header holding
  // exactly one piece of text. Reading BATCH_SET as well as SET_CELL is the
  // point -- see collectSheetCells. TASKS.md #164.
  for (const cell of collectSheetCells(actions, header.sheetName)) {
    if (cell.row >= header.row || !cell.label) continue;
    rows.set(cell.row, (rows.get(cell.row) ?? 0) + 1);
  }

  const singles = [...rows.entries()].filter(([, count]) => count === 1).map(([row]) => row);
  if (singles.length === 0) return null;
  return Math.min(...singles);
}

interface SheetCell {
  row: number;
  col: number;
  /** Non-empty, non-numeric text -- i.e. something that reads as a caption. */
  label: string | null;
  /** True when the cell carries a formula or a numeric value. */
  isValue: boolean;
}

/**
 * Every cell this batch writes to one sheet, whichever action shape carried it.
 *
 * `findTitleRow` read only SET_CELL, so when TASKS.md #157 made BATCH_SET
 * actually work the Executor started writing `A1="Dashboard"` inside a
 * BATCH_SET and the dashboard title quietly stopped being styled. That is the
 * same lesson as #143 (ADD_ROW) and #157 (BATCH_SET) in a fourth place, so the
 * shape-handling is centralised here rather than repeated a fourth time.
 */
export function collectSheetCells(
  actions: SheetActionPayload[],
  sheetName: string,
): SheetCell[] {
  const key = sheetName.trim().toLowerCase();
  const cells: SheetCell[] = [];

  const push = (row: number, col: number, value: unknown, formula: unknown) => {
    const text = typeof value === 'string' ? value.trim() : '';
    const label = text !== '' && Number.isNaN(Number(text)) ? text : null;
    const isValue =
      (typeof formula === 'string' && formula.trim() !== '') ||
      typeof value === 'number' ||
      (text !== '' && !Number.isNaN(Number(text)));
    if (!label && !isValue) return;
    cells.push({ row, col, label, isValue });
  };

  for (const action of actions) {
    if (String(action.sheetName ?? '').trim().toLowerCase() !== key) continue;

    if (action.type === 'SET_CELL' || action.type === 'SET_FORMULA') {
      if (typeof action.row !== 'number' || typeof action.col !== 'number') continue;
      push(action.row, action.col, action.value, action.formula);
      continue;
    }

    if (action.type === 'BATCH_SET' && Array.isArray(action.operations)) {
      for (const op of action.operations) {
        if (typeof op?.address !== 'string') continue;
        const parsed = parseA1Address(op.address);
        if (!parsed) continue;
        push(parsed.row, parsed.col, op.value, op.formula);
      }
    }
  }

  return cells;
}

/** The veto runs first and wins over money/count words. TASKS.md #158. */
function numberFormatForLabel(label: string, currencyFormat: string): string | null {
  if (NON_NUMERIC_HEADER_PATTERN.test(label)) return null;
  if (COUNT_HEADER_PATTERN.test(label)) return '#,##0';
  if (CURRENCY_HEADER_PATTERN.test(label)) return currencyFormat;
  return null;
}

/**
 * Rows of number-format runway below `run` before the next table on that sheet.
 */
function runwayFor(run: HeaderRun, runs: HeaderRun[]): number {
  const key = run.sheetName.toLowerCase();
  let nextRow = Infinity;
  for (const other of runs) {
    if (other === run) continue;
    if (other.sheetName.toLowerCase() !== key) continue;
    if (other.row > run.row && other.row < nextRow) nextRow = other.row;
  }
  if (!Number.isFinite(nextRow)) return FORMAT_RUNWAY_ROWS;
  // Stop on the row before the next header; never emit a zero/negative count.
  return Math.max(1, Math.min(FORMAT_RUNWAY_ROWS, nextRow - run.row - 1));
}

/** Minimum label/value pairs before a row reads as a KPI band. */
const MIN_KPI_PAIRS = 2;
const KPI_FILL = '#EDF3FA';
const KPI_LABEL_FONT = '#595959';

/**
 * Style the summary tiles above a sheet's first table.
 *
 * A live Main sheet wrote `A2="Total Amount" B2==SUM(B5:B16)`,
 * `C2="Paid Amount" D2==SUM(C5:C16)`, `E2="Pending Amount" F2==SUM(D5:D16)` and
 * got **no formatting at all** -- #144 had (correctly) taught `detectHeaderRuns`
 * to reject this row as a table header, because its labels are not contiguous,
 * and nothing else claimed it. So the tables below looked finished while the
 * dashboard itself was raw unstyled cells with unformatted totals.
 *
 * Styling it as paired tiles also resolves a genuine reading hazard that is NOT
 * fixable here: the values sit in B/D/F while the table below has "Pending
 * Amount" as its column D, so the Paid total appears directly above the Pending
 * column. This pass is layout-preserving by contract (see the header comment) --
 * it must not move a cell, because the planner's anchors are computed from those
 * positions. Painting each label and its value as one bordered tile makes the
 * pairing explicit instead, so the column beneath is no longer read as a
 * caption for it. TASKS.md #164.
 */
function styleKpiBand(
  actions: SheetActionPayload[],
  header: HeaderRun,
  currencyFormat: string,
): SheetActionPayload[] {
  const cells = collectSheetCells(actions, header.sheetName).filter((c) => c.row < header.row);
  if (cells.length === 0) return [];

  const at = new Map<string, SheetCell>();
  for (const cell of cells) at.set(`${cell.row}#${cell.col}`, cell);

  const pairs: { label: SheetCell; value: SheetCell }[] = [];
  for (const cell of cells) {
    if (!cell.label) continue;
    // The value sits to the right (A2/B2 tiles) or directly below (label over
    // value); both shapes show up depending on how the Planner lays the band out.
    const right = at.get(`${cell.row}#${cell.col + 1}`);
    const below = at.get(`${cell.row + 1}#${cell.col}`);
    const value =
      right?.isValue && !right.label ? right : below?.isValue && !below.label ? below : null;
    if (!value) continue;
    // A cell that is itself somebody's value is not also a label.
    if (at.get(`${cell.row}#${cell.col - 1}`)?.label) continue;
    pairs.push({ label: cell, value });
  }

  if (pairs.length < MIN_KPI_PAIRS) return [];

  const styling: SheetActionPayload[] = [];
  for (const { label, value } of pairs) {
    styling.push({
      type: 'FORMAT_RANGE',
      sheetName: header.sheetName,
      row: label.row,
      col: label.col,
      rowCount: 1,
      colCount: 1,
      format: {
        bold: true,
        fontColor: KPI_LABEL_FONT,
        fillColor: KPI_FILL,
        verticalAlignment: 'middle',
        borders: 'all',
      },
    });

    const numberFormat = numberFormatForLabel(label.label as string, currencyFormat);
    styling.push({
      type: 'FORMAT_RANGE',
      sheetName: header.sheetName,
      row: value.row,
      col: value.col,
      rowCount: 1,
      colCount: 1,
      format: {
        bold: true,
        fontSize: 12,
        fontColor: TITLE_FONT,
        fillColor: KPI_FILL,
        horizontalAlignment: 'left',
        verticalAlignment: 'middle',
        borders: 'all',
        ...(numberFormat ? { numberFormat } : {}),
      },
    });
  }

  return styling;
}
