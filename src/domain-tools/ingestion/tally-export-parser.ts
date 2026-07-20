import { SourceRef } from '../types/domain-tool.types';

export interface NormalizedTallyRow {
  voucherType: string;
  voucherNumber: string;
  date: string;
  ledgerName: string;
  amount: number;
  sourceRowRef: SourceRef;
}

/**
 * STUB — Tally XML/CSV normalization requires a CA-reviewed fixture set.
 */
export function parseTallyExport(_rawExport: Buffer | string): NormalizedTallyRow[] {
  throw new Error('Not implemented.');
}
