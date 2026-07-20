import { SourceRef } from '../types/domain-tool.types';

export interface Normalized26asRow {
  pan: string;
  deducteeName: string;
  amountPaid: number;
  tdsDeducted: number;
  section: string;
  sourceRowRef: SourceRef;
}

/**
 * STUB — Form 26AS PDF/text normalization requires a CA-reviewed fixture set.
 */
export function parseForm26as(_rawExport: Buffer | string): Normalized26asRow[] {
  throw new Error('Not implemented.');
}
