import { ColumnMapping, NormalizedInvoiceRow } from '../types/domain-tool.types';
import { parseInvoiceGrid, parseRowsFromObjects } from './invoice-grid-parser';

/**
 * Parse Purchase Register sheet/grid into normalized rows.
 */
export function parsePurchaseRegister(
  rawExport: Buffer | string | unknown[][],
  options?: { documentId?: string; columnMapping?: ColumnMapping; headersRow?: number },
): NormalizedInvoiceRow[] {
  const documentId = options?.documentId ?? 'purchase_register';

  if (Array.isArray(rawExport)) {
    return parseInvoiceGrid(rawExport, {
      documentId,
      documentType: 'purchase_register',
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
      documentType: 'purchase_register',
      columnMapping: options?.columnMapping,
      headersRow: options?.headersRow,
    });
  }

  if (parsed?.rows) {
    return parseRowsFromObjects(parsed.rows, 'purchase_register', documentId);
  }

  throw new Error('Unsupported purchase register format.');
}
