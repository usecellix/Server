import { ColumnMapping, NormalizedInvoiceRow } from '../types/domain-tool.types';
import { parseInvoiceGrid, parseRowsFromObjects } from './invoice-grid-parser';

/**
 * Parse GSTR-1 portal export grid into normalized rows.
 * Counterparty = recipient GSTIN (B2B); B2C sections may have blank GSTIN.
 */
export function parseGstr1(
  rawExport: Buffer | string | unknown[][],
  options?: { documentId?: string; columnMapping?: ColumnMapping; headersRow?: number },
): NormalizedInvoiceRow[] {
  const documentId = options?.documentId ?? 'gstr1';

  if (Array.isArray(rawExport)) {
    return parseInvoiceGrid(rawExport, {
      documentId,
      documentType: 'gstr1',
      columnMapping: options?.columnMapping,
      headersRow: options?.headersRow,
      skipEmptyInvoice: true,
    });
  }

  const text = Buffer.isBuffer(rawExport) ? rawExport.toString('utf8') : String(rawExport);
  const parsed = JSON.parse(text) as {
    rows?: Array<Record<string, unknown>>;
  };

  if (Array.isArray(parsed)) {
    return parseInvoiceGrid(parsed as unknown as unknown[][], {
      documentId,
      documentType: 'gstr1',
      columnMapping: options?.columnMapping,
      headersRow: options?.headersRow,
    });
  }

  if (parsed?.rows) {
    return parseRowsFromObjects(parsed.rows, 'gstr1', documentId);
  }

  throw new Error('Unsupported GSTR-1 format.');
}
