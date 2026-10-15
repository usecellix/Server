/**
 * Canonical types for bank reconciliation. Every statement row (any format) and every
 * books-ledger row is normalized into `BankTxn` before matching, so the matcher never
 * knows or cares which format a row came from.
 */

/** Money in / money out from the account holder's viewpoint. */
export type Direction = 'IN' | 'OUT';

export type TxnSource = 'BOOKS' | 'BANK';
export type SourceFormat = 'XLSX' | 'CSV' | 'PDF_DIGITAL' | 'PDF_SCANNED';

export type Channel =
  | 'NEFT'
  | 'RTGS'
  | 'IMPS'
  | 'UPI'
  | 'CHQ'
  | 'ECS'
  | 'NACH'
  | 'ATM'
  | 'POS'
  | 'CASH';

export interface BankTxn {
  /** Stable row id: `<source>:<sheet>:<row>`. */
  id: string;
  source: TxnSource;
  sourceFormat: SourceFormat;
  /** ISO yyyy-mm-dd. */
  txnDate: string;
  valueDate?: string;
  /** Always positive. */
  amount: number;
  direction: Direction;
  instrumentNo?: string;
  referenceNo?: string;
  narration: string;
  counterparty?: string;
  channel?: Channel;
  voucherType?: string;
  voucherNo?: string;
  runningBalance?: number;
  /** 0..1 — below 0.9 means the row is listed under "Please verify". */
  parseConfidence: number;
  /**
   * Where the row came from. `external` rows come from an attached file rather than a sheet
   * in the workbook, so there is no cell to jump to; `ref` is the place in the file ("p3 l12").
   */
  sourceLocation: { sheet: string; row: number; ref?: string; external?: boolean };
}

export interface ParsedTable {
  txns: BankTxn[];
  /** Opening balance (debit/asset-positive for books; credit-positive for bank). */
  openingBalance: number | null;
  closingBalance: number | null;
  headerRow: number;
  /** Text above the header row (account number, bank name, period...). */
  preambleText: string[];
  warnings: string[];
}

export type BooksReason =
  | 'cheque_issued_not_cleared'
  | 'deposit_in_transit'
  | 'stale_cheque'
  | 'post_dated_cheque'
  | 'amount_mismatch'
  | 'date_mismatch_outside_window'
  | 'duplicate_in_books'
  | 'genuinely_missing';

export type BankReason =
  | 'bank_charges_not_booked'
  | 'interest_credited_not_booked'
  | 'tds_on_interest_not_booked'
  | 'direct_credit_not_booked'
  | 'direct_debit_not_booked'
  | 'return_credit_not_booked'
  | 'cheque_bounce_not_booked'
  | 'unidentified_transaction'
  | 'possible_duplicate_in_bank'
  | 'amount_mismatch'
  | 'date_mismatch_outside_window';

export type MatchKind = 'ONE_TO_ONE' | 'GROUP' | 'REVERSAL';

export interface BankMatch {
  id: string;
  kind: MatchKind;
  /** 1 = reference, 2 = amount+date, 3 = narration, 4 = tie-break, 5 = group, 7 = reversal. */
  pass: number;
  confidence: number;
  books: BankTxn[];
  bank: BankTxn[];
  /** For REVERSAL pairs, which side the pair lives on. */
  side?: TxnSource;
}

export interface UnmatchedItem<R extends string = string> {
  txn: BankTxn;
  reason: R;
  explanation: string;
}

export interface BankMatchSettings {
  /** Max days bank may post after books (cheque clearing). */
  clearingWindowDays: number;
  /** Max days bank may post BEFORE books. */
  earlyWindowDays: number;
  /** Tight window for Pass 2 (±). */
  tightWindowDays: number;
  amountTolerance: number;
  narrationThreshold: number;
  maxGroupSize: number;
  enableGroupMatch: boolean;
}

export const DEFAULT_BANK_MATCH_SETTINGS: BankMatchSettings = {
  clearingWindowDays: 15,
  earlyWindowDays: 3,
  tightWindowDays: 3,
  amountTolerance: 0,
  narrationThreshold: 0.5,
  maxGroupSize: 6,
  enableGroupMatch: true,
};

export interface BankMatchResult {
  matches: BankMatch[];
  unmatchedBooks: UnmatchedItem<BooksReason>[];
  unmatchedBank: UnmatchedItem<BankReason>[];
}

export interface ValidationIssue {
  code: string;
  message: string;
  /** 1-based sheet row the issue is localized to, when known. */
  row?: number;
}

export interface ValidationResult {
  /** False when at least one blocking error is present. */
  ok: boolean;
  errors: ValidationIssue[];
  warnings: ValidationIssue[];
}

export interface BrsLine {
  label: string;
  amount: number | null;
  kind: 'balance' | 'add' | 'less' | 'total' | 'difference';
}

export interface BrsResult {
  periodEnd: string;
  booksBalance: number | null;
  bankBalance: number | null;
  /** Items on the bank side not yet in books (bank-only as at period end). */
  bankOnlyIn: number;
  bankOnlyOut: number;
  /** Items on the books side not yet in the bank (as at period end). */
  booksOnlyIn: number;
  booksOnlyOut: number;
  adjustedBooks: number | null;
  adjustedBank: number | null;
  difference: number | null;
  tiedOut: boolean;
  lines: BrsLine[];
  notes: string[];
}
