import { DomainTool, SourceRef } from '../types/domain-tool.types';

export interface BankTxnRow {
  date: string;
  description: string;
  amount: number;
  type: 'credit' | 'debit';
  sourceRowRef: SourceRef;
}

export interface BookTxnRow {
  date: string;
  description: string;
  amount: number;
  type: 'credit' | 'debit';
  sourceRowRef: SourceRef;
}

export interface BankReconInput {
  bankStatement: BankTxnRow[];
  books: BookTxnRow[];
  amountTolerance: number;
  dateWindowDays: number;
}

export interface BankMatchedPair {
  bankRow: BankTxnRow;
  booksRow: BookTxnRow;
  confidence: number;
}

export interface BankReconOutput {
  matched: BankMatchedPair[];
  unmatchedBank: BankTxnRow[];
  unmatchedBooks: BookTxnRow[];
}

/**
 * STUB — bank reconciliation matching requires CA-reviewed rules before production use.
 */
export const bankRecon: DomainTool<BankReconInput, BankReconOutput> = (_input) => {
  throw new Error(
    'Not implemented — requires CA-reviewed bank reconciliation spec before production use.',
  );
};
