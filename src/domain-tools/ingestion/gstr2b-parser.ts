import { NormalizedInvoiceRow } from '../types/domain-tool.types';

/**
 * STUB — real implementation needs the actual GSTR-2B JSON/Excel schema
 * from a CA-reviewed fixture set.
 */
export function parseGstr2b(_rawExport: Buffer | string): NormalizedInvoiceRow[] {
  throw new Error('Not implemented.');
}
