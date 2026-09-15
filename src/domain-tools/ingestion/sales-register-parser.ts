import { ColumnMapping, NormalizedInvoiceRow } from '../types/domain-tool.types';
import { parseInvoiceGrid, parseRowsFromObjects } from './invoice-grid-parser';

/**
 * Parse Sales Register sheet/grid into normalized rows.
 * Counterparty GSTIN = recipient (buyer); blank GSTIN is valid for B2C.
 */
export function parseSalesRegister(
  rawExport: Buffer | string | unknown[][],
  options?: { documentId?: string; columnMapping?: ColumnMapping; headersRow?: number },
): NormalizedInvoiceRow[] {
  const documentId = options?.documentId ?? 'sales_register';

  if (Array.isArray(rawExport)) {
    return parseInvoiceGrid(rawExport, {
      documentId,
      documentType: 'sales_register',
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
      documentType: 'sales_register',
      columnMapping: options?.columnMapping,
      headersRow: options?.headersRow,
    });
  }

  if (parsed?.rows) {
    return parseRowsFromObjects(parsed.rows, 'sales_register', documentId);
  }

  throw new Error('Unsupported sales register format.');
}
