import { NormalizedInvoiceRow } from '../types/domain-tool.types';
import { parseInvoiceGrid, parseRowsFromObjects } from './invoice-grid-parser';
import { ColumnMapping } from '../types/domain-tool.types';

/**
 * Parse GSTR-2B export.
 * Accepts JSON string of the synthetic fixture shape, a raw 2D grid JSON, or Buffer of same.
 */
export function parseGstr2b(
  rawExport: Buffer | string | unknown[][],
  options?: { documentId?: string; columnMapping?: ColumnMapping },
): NormalizedInvoiceRow[] {
  return parsePortalExport(rawExport, 'gstr2b', options?.documentId ?? 'gstr2b', options);
}

export function parseGstr2a(
  rawExport: Buffer | string | unknown[][],
  options?: { documentId?: string; columnMapping?: ColumnMapping },
): NormalizedInvoiceRow[] {
  return parsePortalExport(rawExport, 'gstr2a', options?.documentId ?? 'gstr2a', options);
}

function parsePortalExport(
  rawExport: Buffer | string | unknown[][],
  documentType: 'gstr2b' | 'gstr2a',
  documentId: string,
  options?: { columnMapping?: ColumnMapping },
): NormalizedInvoiceRow[] {
  if (Array.isArray(rawExport)) {
    return parseInvoiceGrid(rawExport, {
      documentId,
      documentType,
      columnMapping: options?.columnMapping,
    });
  }

  const text = Buffer.isBuffer(rawExport) ? rawExport.toString('utf8') : String(rawExport);
  const parsed = JSON.parse(text) as
    | unknown[][]
    | { rows?: Array<Record<string, unknown>>; headers?: unknown[] };

  if (Array.isArray(parsed)) {
    // Could be grid or array of objects
    if (parsed.length && Array.isArray(parsed[0])) {
      return parseInvoiceGrid(parsed as unknown[][], {
        documentId,
        documentType,
        columnMapping: options?.columnMapping,
      });
    }
  }

  if (parsed && typeof parsed === 'object' && Array.isArray((parsed as { rows?: unknown }).rows)) {
    const body = parsed as { rows: Array<Record<string, unknown>>; headers?: unknown[] };
    if (body.headers?.length) {
      const grid: unknown[][] = [
        body.headers,
        ...body.rows.map((r) => [
          r.gstin,
          r.invoiceNumber ?? r.invoiceNo,
          r.invoiceDate,
          r.taxableValue ?? r.taxableAmt,
          r.igst,
          r.cgst,
          r.sgst,
          r.documentType,
        ]),
      ];
      return parseInvoiceGrid(grid, {
        documentId,
        documentType,
        columnMapping: options?.columnMapping ?? {
          gstin: 0,
          invoiceNo: 1,
          invoiceDate: 2,
          taxableAmt: 3,
          igst: 4,
          cgst: 5,
          sgst: 6,
          documentType: 7,
        },
      });
    }
    return parseRowsFromObjects(body.rows, documentType, documentId);
  }

  throw new Error('Unsupported GSTR portal export format.');
}
