import { SourceRef } from '../types/domain-tool.types';

export interface NormalizedBankStatementRow {
  date: string;
  description: string;
  amount: number;
  type: 'credit' | 'debit';
  balance?: number;
  sourceRowRef: SourceRef;
}

/**
 * STUB — bank statement PDF/CSV normalization requires a CA-reviewed fixture set.
 */
export function parseBankStatement(
  _rawExport: Buffer | string,
): NormalizedBankStatementRow[] {
  throw new Error('Not implemented.');
}
