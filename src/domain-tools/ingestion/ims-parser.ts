import { ColumnMapping, NormalizedInvoiceRow } from '../types/domain-tool.types';
import { parseInvoiceGrid, parseRowsFromObjects } from './invoice-grid-parser';

/**
 * Parse IMS Excel export into normalized rows with imsAction populated.
 */
export function parseImsExport(
  rawExport: Buffer | string | unknown[][],
  options?: { documentId?: string; columnMapping?: ColumnMapping; headersRow?: number },
): NormalizedInvoiceRow[] {
  const documentId = options?.documentId ?? 'ims';

  if (Array.isArray(rawExport)) {
    return parseInvoiceGrid(rawExport, {
      documentId,
      documentType: 'ims',
      columnMapping: options?.columnMapping,
      headersRow: options?.headersRow,
    });
  }

  const text = Buffer.isBuffer(rawExport) ? rawExport.toString('utf8') : String(rawExport);
  const parsed = JSON.parse(text) as {
    rows?: Array<Record<string, unknown>>;
    headers?: unknown[];
  };

  if (Array.isArray(parsed)) {
    return parseInvoiceGrid(parsed as unknown as unknown[][], {
      documentId,
      documentType: 'ims',
      columnMapping: options?.columnMapping,
      headersRow: options?.headersRow,
    });
  }

  if (parsed?.rows) {
    return parseRowsFromObjects(parsed.rows, 'ims', documentId);
  }

  throw new Error('Unsupported IMS export format.');
}
