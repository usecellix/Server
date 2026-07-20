import { DomainTool, SourceRef } from '../types/domain-tool.types';

export interface TdsLedgerRow {
  pan: string;
  deducteeName: string;
  amountPaid: number;
  tdsDeducted: number;
  section: string;
  sourceRowRef: SourceRef;
}

export interface Form26asRow {
  pan: string;
  deducteeName: string;
  amountPaid: number;
  tdsDeducted: number;
  section: string;
  sourceRowRef: SourceRef;
}

export interface Tds26asMatchInput {
  books: TdsLedgerRow[];
  form26as: Form26asRow[];
  amountTolerance: number;
}

export interface TdsMatchedPair {
  booksRow: TdsLedgerRow;
  form26asRow: Form26asRow;
  confidence: number;
}

export interface Tds26asMatchOutput {
  matched: TdsMatchedPair[];
  missingIn26as: TdsLedgerRow[];
  missingInBooks: Form26asRow[];
}

/**
 * STUB — 26AS vs books matching requires CA-reviewed rules before production use.
 */
export const tds26asMatch: DomainTool<Tds26asMatchInput, Tds26asMatchOutput> = (_input) => {
  throw new Error(
    'Not implemented — requires CA-reviewed 26AS matching spec before production use.',
  );
};
