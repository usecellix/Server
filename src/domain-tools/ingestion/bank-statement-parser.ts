import { SourceRef } from '../types/domain-tool.types';
import { RawTable, RawTableCell, RawTableRow } from './raw-table.types';
import {
  inferNumericDateOrder,
  looksLikeDate,
  NumericDateOrder,
  paiseToAmount,
  ParsedAmount,
  parseStatementAmount,
  parseStatementDate,
  toIsoDate,
  tokenKind,
} from './statement-values';

/**
 * Bank statement normalization (Root/ATTACHMENT_EXTRACTION_PLAN.md §6).
 *
 * Deterministic by rule: nothing here guesses a value. A row that cannot be
 * read cleanly is kept and flagged, never dropped, and every line of the file
 * after the header is accounted for as a transaction, a continuation line, a
 * recognised non-transaction line, or a `skipped` entry the caller can report.
 */

export interface NormalizedBankStatementRow {
  /** ISO yyyy-mm-dd. Empty when the date could not be read (the row is flagged). */
  date: string;
  valueDate?: string;
  description: string;
  /** Cheque number when the statement has one, otherwise the bank's transaction ID. */
  refNo?: string;
  amount: number;
  type: 'credit' | 'debit';
  /** Signed: negative when the account is overdrawn. */
  balance?: number;
  /** Reasons this row needs a human look. Empty when the row read cleanly. */
  flags: string[];
  sourceRowRef: SourceRef;
}

export type BankStatementColumnRole =
  | 'date'
  | 'valueDate'
  | 'description'
  | 'refNo'
  | 'txnId'
  | 'debit'
  | 'credit'
  | 'amount'
  | 'drcr'
  | 'balance'
  | 'other';

export type BankStatementAmountLayout = 'debit_credit' | 'amount_drcr' | 'signed_amount';

export interface BankStatementSkippedRow {
  ref: string;
  text: string;
  /** `after_total`: looks like a transaction but sits below the statement's total line. */
  reason: 'unrecognised' | 'after_total';
}

export interface BankStatementParseResult {
  rows: NormalizedBankStatementRow[];
  amountLayout: BankStatementAmountLayout;
  hasBalanceColumn: boolean;
  /** Stated on the statement itself, when it prints one. */
  openingBalance?: number;
  closingBalance?: number;
  statedTotalDebit?: number;
  statedTotalCredit?: number;
  /** All but the last four digits replaced, e.g. "XXXXXXXXXX1234". */
  accountNumberMasked?: string;
  skipped: BankStatementSkippedRow[];
  /** `ref` of the header row, for diagnostics. */
  headerRef: string;
}

export type BankStatementParseErrorCode = 'empty' | 'no_header' | 'no_amount_columns' | 'no_transactions';

export class BankStatementParseError extends Error {
  constructor(
    readonly code: BankStatementParseErrorCode,
    message: string,
  ) {
    super(message);
    this.name = 'BankStatementParseError';
  }
}

export interface BankStatementParseOptions {
  /** Recorded on every row's `sourceRowRef`. Defaults to the file name. */
  documentId?: string;
}

// ── Header recognition ──────────────────────────────────────────────────────

// Matched exactly after normalization. Loose "contains" matching is what turns
// an account-details line such as "Date of Issue ... Available Balance" into a
// false header.
const HEADER_ALIASES: Record<Exclude<BankStatementColumnRole, 'other'>, string[]> = {
  date: [
    'date', 'txn date', 'tran date', 'trans date', 'transaction date', 'posting date',
    'post date', 'posted date', 'book date', 'booking date', 'entry date', 'txn dt', 'tran dt',
  ],
  valueDate: ['value date', 'val date', 'value dt', 'val dt'],
  description: [
    'particulars', 'narration', 'description', 'details', 'remarks', 'narrative',
    'transaction details', 'transaction particulars', 'transaction description',
    'transaction remarks', 'transaction narration', 'txn details', 'txn description',
  ],
  refNo: [
    'cheque', 'chq', 'cheque no', 'chq no', 'cheque number', 'chq number', 'cheque details',
    'chq ref no', 'chq ref number', 'cheque ref no', 'ref no', 'ref number', 'reference',
    'reference no', 'reference number', 'instrument no', 'instrument number', 'instr no',
    'ref no cheque no', 'cheque no ref no', 'utr', 'utr no',
  ],
  txnId: ['tran id', 'transaction id', 'txn id', 'txn no', 'transaction no', 'transaction number'],
  debit: [
    'withdrawal', 'withdrawals', 'withdrawal amt', 'withdrawal amount', 'withdrawal dr',
    'withdrawals dr', 'debit', 'debits', 'debit amt', 'debit amount', 'debit dr', 'dr',
    'dr amount', 'dr amt', 'paid out', 'money out',
  ],
  credit: [
    'deposit', 'deposits', 'deposit amt', 'deposit amount', 'deposit cr', 'deposits cr',
    'credit', 'credits', 'credit amt', 'credit amount', 'credit cr', 'cr', 'cr amount',
    'cr amt', 'paid in', 'money in',
  ],
  amount: ['amount', 'amt', 'transaction amount', 'txn amount', 'tran amount'],
  drcr: [
    'dr cr', 'cr dr', 'debit credit', 'dr or cr', 'type', 'tran type', 'txn type',
    'transaction type',
  ],
  balance: [
    'balance', 'bal', 'closing balance', 'running balance', 'balance amount', 'balance amt',
    'available balance', 'line balance', 'ledger balance',
  ],
};

const ALIAS_TO_ROLE = new Map<string, BankStatementColumnRole>();
for (const [role, aliases] of Object.entries(HEADER_ALIASES)) {
  for (const alias of aliases) ALIAS_TO_ROLE.set(alias, role as BankStatementColumnRole);
}

const CURRENCY_WORDS = new Set(['inr', 'rs', 'rupees']);
const NUMERIC_ROLES = new Set<BankStatementColumnRole>(['debit', 'credit', 'amount', 'balance']);
const DATE_ROLES = new Set<BankStatementColumnRole>(['date', 'valueDate']);
const DRCR_VALUE_RE = /^(dr|cr|d|c|db|debit|credit)\.?$/i;

function normalizeHeader(text: string): string {
  return String(text ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .split(' ')
    .filter((word) => word && !CURRENCY_WORDS.has(word))
    .join(' ');
}

function classifyHeader(text: string): BankStatementColumnRole | null {
  return ALIAS_TO_ROLE.get(normalizeHeader(text)) ?? null;
}

/** A header needs a date column, a money column, and at least three known columns in all. */
function isHeaderRow(texts: string[]): boolean {
  const roles = new Set<BankStatementColumnRole>();
  for (const text of texts) {
    const role = classifyHeader(text);
    if (role) roles.add(role);
  }
  const hasMoney = [...roles].some((role) => NUMERIC_ROLES.has(role));
  return roles.size >= 3 && roles.has('date') && hasMoney;
}

const HEADER_SCAN_LIMIT = 250;

function findHeaderIndex(rows: RawTableRow[]): number {
  const limit = Math.min(rows.length, HEADER_SCAN_LIMIT);
  for (let i = 0; i < limit; i++) {
    if (isHeaderRow(rows[i].cells.map((cell) => cell.t))) return i;
  }
  return -1;
}

// ── Columns and logical rows ────────────────────────────────────────────────

interface Column {
  role: BankStatementColumnRole;
  x0: number;
  x1: number;
}

interface LogicalRow {
  ref: string;
  page?: number;
  y?: number;
  /** One entry per column. */
  cells: string[];
  /** PDF fragments that fitted no column. */
  stray: string[];
  /** The text pieces as read, before they were placed in columns. */
  fragments: string[];
  /** Every fragment joined, for keyword and repeat detection. */
  text: string;
}

interface HeaderCell {
  text: string;
  x0: number;
  x1: number;
  /** True for a cell that exists only because of a line above or below the main header line. */
  added: boolean;
}

function positioned(cell: RawTableCell): { text: string; x0: number; x1: number } {
  const x0 = Number.isFinite(cell.x0) ? (cell.x0 as number) : 0;
  const x1 = Number.isFinite(cell.x1) ? (cell.x1 as number) : x0;
  return { text: String(cell.t ?? '').trim(), x0, x1: Math.max(x0, x1) };
}

function overlap(a: { x0: number; x1: number }, b: { x0: number; x1: number }): number {
  return Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
}

/** Header labels are often split over two or three lines ("Tran" above "Type", "DR" above "/CR"). */
const HEADER_BLOCK_MAX_GAP = 14;

function buildPositionedColumns(
  rows: RawTableRow[],
  headerIndex: number,
): { columns: Column[]; consumed: Set<number> } {
  const main = rows[headerIndex];
  const cells: HeaderCell[] = main.cells
    .map(positioned)
    .filter((cell) => cell.text)
    .map((cell) => ({ ...cell, added: false }));
  const consumed = new Set<number>([headerIndex]);

  for (const offset of [-2, -1, 1, 2]) {
    const index = headerIndex + offset;
    const row = rows[index];
    if (!row || row.page !== main.page) continue;
    if (main.y === undefined || row.y === undefined) continue;
    if (Math.abs(row.y - main.y) > HEADER_BLOCK_MAX_GAP) continue;
    const fragments = row.cells.map(positioned).filter((cell) => cell.text);
    if (fragments.some((f) => tokenKind(f.text) === 'date' || tokenKind(f.text) === 'money')) continue;

    const above = offset < 0;
    let contributed = false;
    for (const fragment of fragments) {
      const target = cells.find((cell) => overlap(cell, fragment) > 0);
      if (!target) {
        cells.push({ ...fragment, added: true });
        continue;
      }
      const merged = above ? `${fragment.text} ${target.text}` : `${target.text} ${fragment.text}`;
      // A label on the main line is only extended when the result is a known
      // header; otherwise a first data line such as "B/F" would corrupt it.
      if (!target.added && !classifyHeader(merged)) continue;
      target.text = merged;
      target.x0 = Math.min(target.x0, fragment.x0);
      target.x1 = Math.max(target.x1, fragment.x1);
      contributed = contributed || !target.added;
    }
    const addedRecognised = cells.some(
      (cell) => cell.added && classifyHeader(cell.text) && fragments.some((f) => overlap(cell, f) > 0),
    );
    if (contributed || addedRecognised) consumed.add(index);
  }

  const columns = cells
    .map((cell) => ({ role: classifyHeader(cell.text) ?? 'other', x0: cell.x0, x1: cell.x1 }))
    .sort((a, b) => a.x0 - b.x0);
  return { columns, consumed };
}

const MONEY_MAX_EDGE_DISTANCE = 60;
const DATE_MAX_EDGE_DISTANCE = 60;
const INTEGER_MAX_EDGE_DISTANCE = 25;

function nearestBy(
  columns: Column[],
  accept: (column: Column) => boolean,
  distance: (column: Column) => number,
): { index: number; distance: number } | null {
  let best: { index: number; distance: number } | null = null;
  columns.forEach((column, index) => {
    if (!accept(column)) return;
    const d = distance(column);
    if (!best || d < best.distance) best = { index, distance: d };
  });
  return best;
}

function assignPositionedRow(row: RawTableRow, columns: Column[]): LogicalRow {
  const cells = columns.map(() => '');
  const stray: string[] = [];
  const texts: string[] = [];
  const place = (index: number, text: string) => {
    cells[index] = cells[index] ? `${cells[index]} ${text}` : text;
  };
  const isTextColumn = (column: Column) => !NUMERIC_ROLES.has(column.role) && !DATE_ROLES.has(column.role);

  const placeAsText = (fragment: { text: string; x0: number; x1: number }) => {
    const overlapping = nearestBy(columns, isTextColumn, (column) => -overlap(column, fragment));
    if (overlapping && overlapping.distance < 0) return place(overlapping.index, fragment.text);
    const nearest = nearestBy(columns, isTextColumn, (column) =>
      Math.max(column.x0 - fragment.x1, fragment.x0 - column.x1),
    );
    if (nearest) place(nearest.index, fragment.text);
    else stray.push(fragment.text);
  };

  // Amounts are right-aligned, so a wide number starts to the left of its own
  // heading. Matching on the right edge keeps it in the correct column.
  const placeAsMoney = (fragment: { text: string; x0: number; x1: number }, limit: number): boolean => {
    const target = nearestBy(
      columns,
      (column) => NUMERIC_ROLES.has(column.role),
      (column) => Math.abs(column.x1 - fragment.x1),
    );
    if (!target || target.distance > limit) return false;
    place(target.index, fragment.text);
    return true;
  };

  for (const cell of row.cells) {
    const fragment = positioned(cell);
    if (!fragment.text) continue;
    texts.push(fragment.text);
    const kind = tokenKind(fragment.text);

    if (kind === 'date') {
      const target = nearestBy(
        columns,
        (column) => DATE_ROLES.has(column.role),
        (column) => Math.abs(column.x0 - fragment.x0),
      );
      // A second date in the same column means the nearer slot is taken: use the other one.
      const free =
        target && cells[target.index]
          ? nearestBy(
              columns,
              (column) => DATE_ROLES.has(column.role) && !cells[columns.indexOf(column)],
              (column) => Math.abs(column.x0 - fragment.x0),
            )
          : target;
      if (free && free.distance <= DATE_MAX_EDGE_DISTANCE) place(free.index, fragment.text);
      else placeAsText(fragment);
      continue;
    }

    if (kind === 'money') {
      if (!placeAsMoney(fragment, MONEY_MAX_EDGE_DISTANCE)) placeAsText(fragment);
      continue;
    }

    if (kind === 'integer') {
      // A bare integer under a reference-type heading is a cheque or reference
      // number. Anywhere else it is tried as an amount first.
      const underText = columns.some((column) => isTextColumn(column) && overlap(column, fragment) > 0);
      if (underText || !placeAsMoney(fragment, INTEGER_MAX_EDGE_DISTANCE)) placeAsText(fragment);
      continue;
    }

    placeAsText(fragment);
  }

  return { ref: row.ref, page: row.page, y: row.y, cells, stray, fragments: texts, text: texts.join(' ') };
}

function gridRow(row: RawTableRow, columnCount: number): LogicalRow {
  const cells = Array.from({ length: columnCount }, (_, index) =>
    String(row.cells[index]?.t ?? '').trim(),
  );
  const extra = row.cells
    .slice(columnCount)
    .map((cell) => String(cell.t ?? '').trim())
    .filter(Boolean);
  const fragments = [...cells, ...extra].filter(Boolean);
  return { ref: row.ref, cells, stray: [], fragments, text: fragments.join(' ') };
}

// ── Line classification ─────────────────────────────────────────────────────

const OPENING_RE = /^\W*(opening balance|opening bal|balance b\W?f|b\W?f\b|brought forward|balance brought forward)/i;
const CLOSING_RE = /^\W*(closing balance|closing bal|balance c\W?f|c\W?f\b|carried forward|balance carried forward)/i;
const GRAND_TOTAL_RE = /^\W*(grand total|statement total|total for the period)\b/i;
const TOTAL_RE = /^\W*((page|sub)\W*)?totals?\b/i;
const PAGE_TOTAL_RE = /^\W*(page|sub)\W*totals?\b/i;
const END_RE = /end of statement|statement summary|\*{3,}\s*end/i;
const FURNITURE_RE = /^\W*(page\s*\d+(\s*(of|\/)\s*\d+)?|continued\b.*|contd\b.*|this is a computer generated.*)\W*$/i;

function median(values: number[]): number | undefined {
  if (!values.length) return undefined;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

/**
 * Page headers and footers: the same text at the same height on at least half
 * the pages. Height is part of the test so that a merchant line which merely
 * recurs in the descriptions is never mistaken for a footer. `rows` must
 * already exclude anything carrying a date or an amount.
 */
function findPageFurniture(rows: LogicalRow[], pageCount: number): Set<LogicalRow> {
  const furniture = new Set<LogicalRow>();
  if (pageCount < 2) return furniture;
  const threshold = Math.max(2, Math.ceil(pageCount / 2));

  const byText = new Map<string, LogicalRow[]>();
  for (const row of rows) {
    if (row.page === undefined || row.y === undefined || !row.text) continue;
    const key = row.text.replace(/\d+/g, '#');
    const list = byText.get(key) ?? [];
    list.push(row);
    byText.set(key, list);
  }
  for (const list of byText.values()) {
    if (list.length < threshold) continue;
    const centre = median(list.map((row) => row.y as number)) as number;
    const aligned = list.filter((row) => Math.abs((row.y as number) - centre) <= 2);
    if (new Set(aligned.map((row) => row.page)).size >= threshold) {
      for (const row of aligned) furniture.add(row);
    }
  }
  return furniture;
}

function maskAccountNumber(rows: RawTableRow[]): string | undefined {
  for (const row of rows) {
    const text = row.cells.map((cell) => String(cell.t ?? '').trim()).join(' ');
    const match = /\b(?:a\/c|ac|acct|account)\.?\s*(?:no|number|num)\b\.?\s*[:\-]?\s*([0-9Xx*]{6,24})/i.exec(text);
    if (!match) continue;
    const raw = match[1];
    return 'X'.repeat(Math.max(0, raw.length - 4)) + raw.slice(-4);
  }
  return undefined;
}

/** Wrapped PDF descriptions break mid-word, so a part ending or starting on punctuation is rejoined directly. */
function joinDescription(head: string, tail: string, layout: RawTable['layout']): string {
  if (!head) return tail;
  if (!tail) return head;
  if (layout === 'positioned' && (/[^A-Za-z0-9)\]]$/.test(head) || /^[^A-Za-z0-9(\[]/.test(tail))) {
    return head + tail;
  }
  return `${head} ${tail}`;
}

// ── Parser ──────────────────────────────────────────────────────────────────

interface DraftTxn {
  row: LogicalRow;
  description: string;
}

export function parseBankStatement(
  table: RawTable,
  options: BankStatementParseOptions = {},
): BankStatementParseResult {
  const rawRows = table?.rows ?? [];
  if (!rawRows.length) {
    throw new BankStatementParseError('empty', 'The file has no readable rows.');
  }

  const headerIndex = findHeaderIndex(rawRows);
  if (headerIndex < 0) {
    throw new BankStatementParseError(
      'no_header',
      'Could not find the transaction table. A statement needs a header row with a date column and an amount or balance column.',
    );
  }

  let columns: Column[];
  let data: LogicalRow[];
  if (table.layout === 'positioned') {
    const built = buildPositionedColumns(rawRows, headerIndex);
    columns = built.columns;
    data = rawRows
      .map((row, index) => ({ row, index }))
      .filter(({ index }) => index > headerIndex && !built.consumed.has(index))
      .map(({ row }) => assignPositionedRow(row, columns));
  } else {
    columns = rawRows[headerIndex].cells.map((cell) => ({
      role: classifyHeader(cell.t) ?? 'other',
      x0: 0,
      x1: 0,
    }));
    data = rawRows.slice(headerIndex + 1).map((row) => gridRow(row, columns.length));
  }

  const firstIndex = (role: BankStatementColumnRole): number =>
    columns.findIndex((column) => column.role === role);
  const dateIdx = firstIndex('date');
  const hasDate = (row: LogicalRow) => looksLikeDate(row.cells[dateIdx] ?? '');

  // A "type" column is only Dr/Cr when its values say so. Federal Bank's
  // "Tran Type" holds TFR/FT/CASH and must not be read as a direction.
  columns.forEach((column, index) => {
    if (column.role !== 'drcr') return;
    const values = data.filter(hasDate).map((row) => row.cells[index]).filter(Boolean);
    const valid = values.filter((value) => DRCR_VALUE_RE.test(value)).length;
    if (!values.length || valid / values.length < 0.8) column.role = 'other';
  });

  const valueDateIdx = firstIndex('valueDate');
  const descIdx = firstIndex('description');
  const refIdx = firstIndex('refNo');
  const txnIdIdx = firstIndex('txnId');
  const debitIdx = firstIndex('debit');
  const creditIdx = firstIndex('credit');
  const amountIdx = firstIndex('amount');
  const balanceIdx = firstIndex('balance');
  // With separate debit and credit columns, a Dr/Cr column can only be the sign of the balance.
  const drcrIdx = columns.map((column) => column.role).lastIndexOf('drcr');

  let amountLayout: BankStatementAmountLayout;
  if (debitIdx >= 0 && creditIdx >= 0) amountLayout = 'debit_credit';
  else if (amountIdx >= 0) amountLayout = drcrIdx >= 0 ? 'amount_drcr' : 'signed_amount';
  else {
    throw new BankStatementParseError(
      'no_amount_columns',
      'The transaction table has no usable amount columns. It needs debit and credit columns, or a single amount column.',
    );
  }

  const order: NumericDateOrder = inferNumericDateOrder(
    data.filter(hasDate).map((row) => row.cells[dateIdx]),
  );

  const isPositioned = table.layout === 'positioned';

  // Typical distance from a transaction line to the line under it. A footer
  // sits much further below the last transaction than a wrapped description does.
  const gaps: number[] = [];
  if (isPositioned) {
    for (let i = 0; i + 1 < data.length; i++) {
      const here = data[i];
      const next = data[i + 1];
      if (!hasDate(here) || here.page !== next.page) continue;
      if (here.y === undefined || next.y === undefined) continue;
      if (here.y - next.y > 0) gaps.push(here.y - next.y);
    }
  }
  const typicalGap = median(gaps);
  const maxContinuationGap = typicalGap !== undefined ? typicalGap * 1.8 : 20;

  const cell = (row: LogicalRow, index: number): string => (index >= 0 ? row.cells[index] ?? '' : '');
  const amountOf = (row: LogicalRow, index: number): ParsedAmount | null =>
    index >= 0 ? parseStatementAmount(row.cells[index] ?? '') : null;
  const hasMoney = (row: LogicalRow): boolean =>
    [debitIdx, creditIdx, amountIdx, balanceIdx].some((index) => amountOf(row, index) !== null);

  const signedBalance = (row: LogicalRow): number | null => {
    const parsed = amountOf(row, balanceIdx);
    if (!parsed) return null;
    const drcrSaysDebit = amountLayout === 'debit_credit' && /^d/i.test(cell(row, drcrIdx));
    const negative = parsed.negative || parsed.marker === 'dr' || drcrSaysDebit;
    return negative ? -parsed.paise : parsed.paise;
  };
  /** The one figure on an opening or closing line, wherever the bank printed it. */
  const balanceFigure = (row: LogicalRow): number | null => {
    const fromBalance = signedBalance(row);
    if (fromBalance !== null) return fromBalance;
    for (const index of [amountIdx, creditIdx, debitIdx]) {
      const parsed = amountOf(row, index);
      if (parsed) return parsed.negative ? -parsed.paise : parsed.paise;
    }
    return null;
  };
  const onlyDescription = (row: LogicalRow): boolean =>
    descIdx >= 0 &&
    Boolean(row.cells[descIdx]) &&
    row.stray.length === 0 &&
    row.cells.every((value, index) => index === descIdx || !value);

  const pageCount = new Set(data.map((row) => row.page).filter((p) => p !== undefined)).size;
  const furniture = isPositioned
    ? findPageFurniture(
        data.filter((row) => !hasDate(row) && !hasMoney(row)),
        pageCount,
      )
    : new Set<LogicalRow>();

  // Does this statement wrap descriptions onto a second line as a matter of
  // course? When it does, a line directly under a transaction is a wrapped
  // description even if the same text recurs at the same height on other pages
  // (a regular payee). When it does not, such a line is a footer.
  const dated = data.filter(hasDate);
  const wrapped = data.filter(
    (row, index) => hasDate(row) && data[index + 1] !== undefined && onlyDescription(data[index + 1]),
  );
  const wrapsDescriptions = dated.length > 0 && wrapped.length / dated.length >= 0.5;

  const drafts: DraftTxn[] = [];
  const skipped: BankStatementSkippedRow[] = [];
  const clip = (text: string) => (text.length > 160 ? `${text.slice(0, 160)}…` : text);
  let openingPaise: number | null = null;
  let closingPaise: number | null = null;
  let statedDebit: number | null = null;
  let statedCredit: number | null = null;
  let ended = false;
  let lastLine: LogicalRow | null = null;
  const pagesWithContent = new Set<number>();

  for (const row of data) {
    if (!row.text) continue;
    const isDated = hasDate(row);

    if (ended) {
      if (isDated && hasMoney(row)) {
        skipped.push({ ref: row.ref, text: clip(row.text), reason: 'after_total' });
      }
      continue;
    }
    if (!isDated && isHeaderRow(row.fragments)) continue;

    const firstOnPage = row.page !== undefined && !pagesWithContent.has(row.page);
    let continues = false;
    if (!isDated && drafts.length && lastLine && onlyDescription(row)) {
      const samePage = row.page === lastLine.page;
      continues =
        !isPositioned ||
        (samePage &&
          row.y !== undefined &&
          lastLine.y !== undefined &&
          lastLine.y - row.y <= maxContinuationGap) ||
        // A description that wraps across a page break.
        (!samePage && firstOnPage && row.page === (lastLine.page ?? 0) + 1);
    }
    if (furniture.has(row) && !(continues && wrapsDescriptions)) continue;
    if (row.page !== undefined) pagesWithContent.add(row.page);

    const description = cell(row, descIdx);
    const noMovement = amountOf(row, debitIdx) === null && amountOf(row, creditIdx) === null && amountOf(row, amountIdx) === null;

    if (isDated) {
      if (noMovement && OPENING_RE.test(description)) {
        openingPaise = balanceFigure(row) ?? openingPaise;
        continue;
      }
      if (noMovement && CLOSING_RE.test(description)) {
        closingPaise = balanceFigure(row) ?? closingPaise;
        continue;
      }
      drafts.push({ row, description });
      lastLine = row;
      continue;
    }

    if (hasMoney(row)) {
      if (OPENING_RE.test(row.text)) {
        openingPaise = balanceFigure(row) ?? openingPaise;
        continue;
      }
      if (CLOSING_RE.test(row.text)) {
        closingPaise = balanceFigure(row) ?? closingPaise;
        continue;
      }
      if (GRAND_TOTAL_RE.test(row.text) || (TOTAL_RE.test(row.text) && !PAGE_TOTAL_RE.test(row.text))) {
        statedDebit = amountOf(row, debitIdx)?.paise ?? statedDebit;
        statedCredit = amountOf(row, creditIdx)?.paise ?? statedCredit;
        if (GRAND_TOTAL_RE.test(row.text)) ended = true;
        continue;
      }
      if (PAGE_TOTAL_RE.test(row.text)) continue;
    }
    if (END_RE.test(row.text)) {
      ended = true;
      continue;
    }

    if (continues) {
      const draft = drafts[drafts.length - 1];
      draft.description = joinDescription(draft.description, description, table.layout);
      lastLine = row;
      continue;
    }

    if (FURNITURE_RE.test(row.text)) continue;
    skipped.push({ ref: row.ref, text: clip(row.text), reason: 'unrecognised' });
  }

  if (!drafts.length) {
    throw new BankStatementParseError(
      'no_transactions',
      'Found the transaction table header but no transaction rows under it.',
    );
  }

  const documentId = options.documentId ?? table.fileName ?? 'bank-statement';
  const rows = drafts.map(({ row, description }): NormalizedBankStatementRow => {
    const flags: string[] = [];
    const unread = (index: number, label: string) => {
      const text = cell(row, index);
      if (text && parseStatementAmount(text) === null && !/^[-–—.\s]*$/.test(text)) {
        flags.push(`${label} not read: "${clip(text)}"`);
      }
    };

    const date = parseStatementDate(cell(row, dateIdx), order);
    if (!date) flags.push(`Date not read: "${clip(cell(row, dateIdx))}"`);
    const valueDate = valueDateIdx >= 0 ? parseStatementDate(cell(row, valueDateIdx), order) : null;

    let debitPaise = 0;
    let creditPaise = 0;
    if (amountLayout === 'debit_credit') {
      const debit = amountOf(row, debitIdx);
      const credit = amountOf(row, creditIdx);
      unread(debitIdx, 'Debit');
      unread(creditIdx, 'Credit');
      debitPaise = debit?.paise ?? 0;
      creditPaise = credit?.paise ?? 0;
      if (!debit && !credit) flags.push('No amount on this row');
      if (debitPaise > 0 && creditPaise > 0) flags.push('Both a debit and a credit on this row');
    } else {
      const amount = amountOf(row, amountIdx);
      unread(amountIdx, 'Amount');
      if (!amount) {
        flags.push('No amount on this row');
      } else {
        const drcr = amountLayout === 'amount_drcr' ? cell(row, drcrIdx) : '';
        const isDebit = amount.marker
          ? amount.marker === 'dr'
          : drcr
            ? /^d/i.test(drcr)
            : amount.negative;
        if (amountLayout === 'amount_drcr' && !amount.marker && !drcr) {
          flags.push('No Dr/Cr marker on this row');
        }
        if (isDebit) debitPaise = amount.paise;
        else creditPaise = amount.paise;
      }
    }

    unread(balanceIdx, 'Balance');
    const balancePaise = signedBalance(row);
    if (row.stray.length) flags.push(`Text not placed in a column: "${clip(row.stray.join(' '))}"`);

    const net = creditPaise - debitPaise;
    const refNo = cell(row, refIdx) || cell(row, txnIdIdx);
    return {
      date: date ? toIsoDate(date) : '',
      ...(valueDate ? { valueDate: toIsoDate(valueDate) } : {}),
      description: description.trim(),
      ...(refNo ? { refNo } : {}),
      amount: paiseToAmount(Math.abs(net)),
      type: creditPaise > 0 && net >= 0 ? 'credit' : 'debit',
      ...(balancePaise !== null ? { balance: paiseToAmount(balancePaise) } : {}),
      flags,
      sourceRowRef: { documentType: 'bank_statement', documentId, rowOrLine: row.ref },
    };
  });

  const accountNumberMasked = maskAccountNumber(rawRows.slice(0, headerIndex));
  return {
    rows,
    amountLayout,
    hasBalanceColumn: balanceIdx >= 0 && rows.some((row) => row.balance !== undefined),
    ...(openingPaise !== null ? { openingBalance: paiseToAmount(openingPaise) } : {}),
    ...(closingPaise !== null ? { closingBalance: paiseToAmount(closingPaise) } : {}),
    ...(statedDebit !== null ? { statedTotalDebit: paiseToAmount(statedDebit) } : {}),
    ...(statedCredit !== null ? { statedTotalCredit: paiseToAmount(statedCredit) } : {}),
    ...(accountNumberMasked ? { accountNumberMasked } : {}),
    skipped,
    headerRef: rawRows[headerIndex].ref,
  };
}
