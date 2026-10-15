import { SheetActionPayload } from '../excel-ai/types/sheet-actions.types';
import {
  buildReconSheetStylingActions,
  StyledSheetBuilder,
} from '../excel-ai/utils/styled-recon-sheet.util';
import {
  BankMatch,
  BankMatchResult,
  BankReason,
  BankTxn,
  BooksReason,
  BrsResult,
  UnmatchedItem,
} from './types';
import { BankReconLayoutDto } from './bank-recon.dto';

export const BOOKS_REASON_LABEL: Record<BooksReason, string> = {
  cheque_issued_not_cleared: 'Cheques issued, not yet presented',
  deposit_in_transit: 'Deposits in transit',
  stale_cheque: 'Stale cheques (older than 3 months)',
  post_dated_cheque: 'Post-dated entries (after period end)',
  amount_mismatch: 'Amount differs from bank',
  date_mismatch_outside_window: 'Date differs from bank beyond clearing window',
  duplicate_in_books: 'Possible duplicate postings in books',
  genuinely_missing: 'In books, not in bank — no explanation found',
};

export const BANK_REASON_LABEL: Record<BankReason, string> = {
  bank_charges_not_booked: 'Bank charges not booked',
  interest_credited_not_booked: 'Interest credited, not booked',
  tds_on_interest_not_booked: 'TDS on interest, not booked',
  direct_credit_not_booked: 'Direct credits, receipt not booked',
  direct_debit_not_booked: 'Direct debits / auto-debits, not booked',
  return_credit_not_booked: 'Returned payments (refund credited), not booked',
  cheque_bounce_not_booked: 'Cheque bounces / returns, not booked',
  unidentified_transaction: 'Unidentified bank entries (CA review)',
  possible_duplicate_in_bank: 'Possible duplicates on the statement',
  amount_mismatch: 'Amount differs from books',
  date_mismatch_outside_window: 'Date differs from books beyond clearing window',
};

const BOOKS_ORDER: BooksReason[] = [
  'genuinely_missing',
  'amount_mismatch',
  'date_mismatch_outside_window',
  'duplicate_in_books',
  'stale_cheque',
  'cheque_issued_not_cleared',
  'deposit_in_transit',
  'post_dated_cheque',
];
const BANK_ORDER: BankReason[] = [
  'amount_mismatch',
  'date_mismatch_outside_window',
  'unidentified_transaction',
  'direct_credit_not_booked',
  'direct_debit_not_booked',
  'return_credit_not_booked',
  'bank_charges_not_booked',
  'interest_credited_not_booked',
  'tds_on_interest_not_booked',
  'cheque_bounce_not_booked',
  'possible_duplicate_in_bank',
];

export interface BankSheetModel {
  result: BankMatchResult;
  brs: BrsResult;
  periodLabel: string;
  periodStart: string;
  periodEnd: string;
  runAt: string;
  relativeTo?: string;
  booksSheet: string;
  bankSheet: string;
}

/** Books-side reasons that are plain timing differences — the "cheques not cleared" view. */
export const TIMING_REASONS: BooksReason[] = ['cheque_issued_not_cleared', 'deposit_in_transit', 'stale_cheque'];

export const NOT_CLEARED_SHEET = 'Bank Recon - Not Cleared';

export function bankSheetName(layout: BankReconLayoutDto): string {
  switch (layout) {
    case 'books_flat':
      return 'Bank Recon - Books Only';
    case 'bank_flat':
      return 'Bank Recon - Bank Only';
    case 'matched':
      return 'Bank Recon - Matched';
    default:
      return 'Bank Recon Summary';
  }
}

const HEADERS = [
  'Date',
  'Voucher / Source',
  'Cheque / Ref No',
  'Narration',
  'Receipt',
  'Payment',
  'Reason',
  'Explanation',
  'Source Row',
];
const WIDTH = HEADERS.length;

function refCell(t: BankTxn): string {
  return [t.instrumentNo, t.referenceNo].filter(Boolean).join(' / ');
}

function sourceCell(t: BankTxn): string {
  const loc = t.sourceLocation;
  // A row from an attached file has no cell in the workbook — point at its place in the file.
  if (loc.external) return `${loc.sheet} · ${loc.ref ?? `row ${loc.row}`}`;
  return `${loc.sheet}!${loc.row}`;
}

function txnRow(t: BankTxn, reason: string, explanation: string): unknown[] {
  const voucher = [t.voucherType, t.voucherNo].filter(Boolean).join(' ') || (t.source === 'BANK' ? 'Bank' : '');
  return [
    t.txnDate,
    voucher,
    refCell(t),
    t.narration,
    t.direction === 'IN' ? t.amount : '',
    t.direction === 'OUT' ? t.amount : '',
    reason,
    explanation,
    sourceCell(t),
  ];
}

function groupBooks(items: UnmatchedItem<BooksReason>[]) {
  return BOOKS_ORDER.map((r) => ({ reason: r, items: items.filter((i) => i.reason === r) })).filter((g) => g.items.length);
}
function groupBank(items: UnmatchedItem<BankReason>[]) {
  return BANK_ORDER.map((r) => ({ reason: r, items: items.filter((i) => i.reason === r) })).filter((g) => g.items.length);
}

/** Rows that only matter for the report — out-of-period bank rows are used for matching but not listed. */
export function reportableBank(model: BankSheetModel): UnmatchedItem<BankReason>[] {
  return model.result.unmatchedBank.filter(
    (u) => u.txn.txnDate >= model.periodStart && u.txn.txnDate <= model.periodEnd,
  );
}

function finish(
  sheetName: string,
  builder: StyledSheetBuilder,
  sheetTitleRow: number,
  relativeTo?: string,
  extra: SheetActionPayload[] = [],
): SheetActionPayload[] {
  const writeHeaders = builder.rows[0].map((c, i) => (i === 0 ? String(c) : `C${i + 1}`));
  const writeRows = builder.rows.slice(1);
  const styling = buildReconSheetStylingActions({
    sheetName,
    columnCount: WIDTH,
    sheetTitleRow,
    sections: builder.sections,
    summaryValueRange: builder.summaryValueRange,
  });
  return [
    { type: 'CREATE_SHEET', sheetName, name: sheetName, relativeTo, position: 'after' },
    { type: 'WRITE_TABLE', sheetName, headers: writeHeaders as string[], rows: writeRows },
    ...styling,
    ...extra,
  ];
}

/** Blank rows between the tables of the summary sheet, so each one reads as its own block. */
const SECTION_GAP_ROWS = 3;

const FOOTER = 'CELLIX | Reconciliation aid. CA to verify.';

export function mapCategorizedToActions(model: BankSheetModel, sheetName: string): SheetActionPayload[] {
  const { result, brs } = model;
  const builder = new StyledSheetBuilder(WIDTH);
  const titleRow = builder.pushRow([`Bank Reconciliation Statement as at ${model.periodEnd}`]);
  builder.pushRow([`Period: ${model.periodLabel}`, `Run: ${model.runAt}`]);
  builder.pushRow([`Books: ${model.booksSheet}`, `Bank: ${model.bankSheet}`]);
  builder.pushRow(['']);

  builder.pushRow(['BANK RECONCILIATION STATEMENT']);
  const brsStart = builder.rows.length;
  builder.beginSummaryBlock();
  for (const line of brs.lines) builder.pushRow([line.label, line.amount ?? 'n/a']);
  builder.endSummaryBlock();
  const brsEnd = builder.rows.length - 1;
  for (const n of brs.notes) builder.pushRow([n]);
  builder.pushRow([
    brs.tiedOut ? 'BRS ties out to 0.00 ✔' : brs.difference === null ? 'BRS could not be tied out (balances missing)' : 'BRS DOES NOT TIE OUT',
  ]);
  builder.pushRow(['']);

  const matched = result.matches.filter((m) => m.kind !== 'REVERSAL');
  builder.pushRow(['SUMMARY']);
  builder.pushRow(['Matched (one-to-one)', matched.filter((m) => m.kind === 'ONE_TO_ONE').length]);
  const carried = matched.filter((m) => m.books.some((b) => b.voucherType === 'Carry-forward')).length;
  if (carried) builder.pushRow(['  of which prior-period items cleared this month', carried]);
  const groups = matched.filter((m) => m.kind === 'GROUP');
  builder.pushRow(['Matched (group)', groups.length]);
  if (groups.length) {
    const gb = groups.reduce((n, m) => n + m.books.length, 0);
    const gk = groups.reduce((n, m) => n + m.bank.length, 0);
    builder.pushRow([`  covering ${gb} ledger entries and ${gk} bank rows (use the 'matched' layout to list them)`, '']);
  }
  const lowMatched = result.matches.flatMap((m) => [...m.books, ...m.bank]).filter((t) => t.parseConfidence < 0.9).length;
  if (lowMatched) builder.pushRow(['Matched rows read with lower confidence', lowMatched]);
  builder.pushRow(['Reversal / bounce pairs', result.matches.filter((m) => m.kind === 'REVERSAL').length]);
  builder.pushRow(['In books, not in bank', result.unmatchedBooks.length]);
  builder.pushRow(['In bank, not in books', reportableBank(model).length]);
  builder.pushRow(['']);

  for (const g of groupBooks(result.unmatchedBooks)) {
    builder.pushSection(
      `${BOOKS_REASON_LABEL[g.reason]} (${g.items.length})`,
      HEADERS,
      g.items.map((i) => txnRow(i.txn, BOOKS_REASON_LABEL[g.reason], i.explanation)),
      SECTION_GAP_ROWS,
    );
  }
  for (const g of groupBank(reportableBank(model))) {
    builder.pushSection(
      `${BANK_REASON_LABEL[g.reason]} (${g.items.length})`,
      HEADERS,
      g.items.map((i) => txnRow(i.txn, BANK_REASON_LABEL[g.reason], i.explanation)),
      SECTION_GAP_ROWS,
    );
  }

  const verify = [...result.unmatchedBooks.map((u) => u.txn), ...reportableBank(model).map((u) => u.txn)].filter(
    (t) => t.parseConfidence < 0.9,
  );
  const verifyAll = verify;
  if (verifyAll.length) {
    builder.pushSection(
      `Please verify — low-confidence rows (${verifyAll.length})`,
      HEADERS,
      verifyAll.map((t) => txnRow(t, 'Low confidence', 'Direction inferred or date borrowed — check against the source sheet.')),
      SECTION_GAP_ROWS,
    );
  }

  builder.pushRow([FOOTER]);

  const amountFmt: SheetActionPayload = {
    type: 'FORMAT_RANGE',
    sheetName,
    range: `B${brsStart + 1}:B${brsEnd + 1}`,
    format: { numberFormat: '#,##0.00' },
  };
  return finish(sheetName, builder, titleRow, model.relativeTo, [amountFmt]);
}

export function mapBooksFlatToActions(model: BankSheetModel, sheetName: string): SheetActionPayload[] {
  const items = model.result.unmatchedBooks;
  if (!items.length) return [];
  const builder = new StyledSheetBuilder(WIDTH);
  const titleRow = builder.pushRow([
    sheetName === NOT_CLEARED_SHEET
      ? `Cheques and deposits not yet cleared by ${model.bankSheet} as at ${model.periodEnd}`
      : `Books entries missing from ${model.bankSheet}`,
  ]);
  builder.pushRow([`Run: ${model.runAt}`, `Rows: ${items.length}`]);
  builder.pushRow(['']);
  builder.pushSection(
    sheetName === NOT_CLEARED_SHEET ? 'Not yet cleared' : 'Missing Rows',
    HEADERS,
    [...items]
      .sort((a, b) => BOOKS_ORDER.indexOf(a.reason) - BOOKS_ORDER.indexOf(b.reason))
      .map((i) => txnRow(i.txn, BOOKS_REASON_LABEL[i.reason], i.explanation)),
  );
  builder.pushRow([FOOTER]);
  return finish(sheetName, builder, titleRow, model.relativeTo);
}

export function mapBankFlatToActions(model: BankSheetModel, sheetName: string): SheetActionPayload[] {
  const items = reportableBank(model);
  if (!items.length) return [];
  const builder = new StyledSheetBuilder(WIDTH);
  const titleRow = builder.pushRow([`Bank statement rows missing from ${model.booksSheet} — entries to post`]);
  builder.pushRow([`Run: ${model.runAt}`, `Rows: ${items.length}`]);
  builder.pushRow(['']);
  builder.pushSection(
    'Entries to post',
    HEADERS,
    [...items]
      .sort((a, b) => BANK_ORDER.indexOf(a.reason) - BANK_ORDER.indexOf(b.reason))
      .map((i) => txnRow(i.txn, BANK_REASON_LABEL[i.reason], i.explanation)),
  );
  builder.pushRow([FOOTER]);
  return finish(sheetName, builder, titleRow, model.relativeTo);
}

export function mapMatchedToActions(model: BankSheetModel, sheetName: string): SheetActionPayload[] {
  const matches = model.result.matches;
  if (!matches.length) return [];
  const headers = ['Date (Books)', 'Date (Bank)', 'Cheque / Ref No', 'Narration', 'Receipt', 'Payment', 'Pass', 'Confidence', 'Source Row'];
  const builder = new StyledSheetBuilder(headers.length);
  const titleRow = builder.pushRow([`Matched entries — ${model.booksSheet} ↔ ${model.bankSheet}`]);
  builder.pushRow([`Run: ${model.runAt}`, `Matches: ${matches.length}`]);
  builder.pushRow(['']);
  const passName: Record<number, string> = {
    1: 'Reference',
    2: 'Amount + date',
    3: 'Narration',
    4: 'Tie-break',
    5: 'Group',
    7: 'Reversal',
  };
  const rows = matches.map((m) => {
    const ref = m.books[0] ?? m.bank[0];
    const amt = (m.books.length ? m.books : m.bank).reduce((s, t) => s + t.amount, 0);
    return [
      m.books.map((b) => b.txnDate).join(', '),
      m.bank.map((k) => k.txnDate).join(', '),
      [...new Set([...m.books, ...m.bank].map(refCell).filter(Boolean))].join(' | '),
      ref.narration,
      ref.direction === 'IN' ? amt : '',
      ref.direction === 'OUT' ? amt : '',
      passName[m.pass] ?? String(m.pass),
      m.confidence,
      [...m.books, ...m.bank].map(sourceCell).join(' ↔ '),
    ];
  });
  builder.pushSection('Matched', headers, rows);
  builder.pushRow([FOOTER]);
  const writeHeaders = builder.rows[0].map((c, i) => (i === 0 ? String(c) : `C${i + 1}`));
  const styling = buildReconSheetStylingActions({
    sheetName,
    columnCount: headers.length,
    sheetTitleRow: titleRow,
    sections: builder.sections,
  });
  return [
    { type: 'CREATE_SHEET', sheetName, name: sheetName, relativeTo: model.relativeTo, position: 'after' },
    { type: 'WRITE_TABLE', sheetName, headers: writeHeaders as string[], rows: builder.rows.slice(1) },
    ...styling,
  ];
}

export function mapBankReconToActions(
  layout: BankReconLayoutDto,
  model: BankSheetModel,
  sheetName: string,
): SheetActionPayload[] {
  switch (layout) {
    case 'books_flat':
      return mapBooksFlatToActions(model, sheetName);
    case 'bank_flat':
      return mapBankFlatToActions(model, sheetName);
    case 'matched':
      return mapMatchedToActions(model, sheetName);
    default:
      return mapCategorizedToActions(model, sheetName);
  }
}

