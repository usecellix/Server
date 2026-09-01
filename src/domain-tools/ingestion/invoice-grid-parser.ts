import {
  ColumnMapping,
  DomainDocumentType,
  NormalizedInvoiceRow,
} from '../types/domain-tool.types';
import { buildNormalizedRow } from '../gst/normalize-invoice';
import {
  cellAt,
  mergeColumnMapping,
  resolveColumnIndex,
} from './portal-file-detector';

export interface GridParseOptions {
  documentId?: string;
  documentType?: DomainDocumentType;
  headersRow?: number;
  columnMapping?: ColumnMapping | null;
  /** When true, skip blank invoice-number rows */
  skipEmptyInvoice?: boolean;
}

function headerRowIndex(options?: GridParseOptions): number {
  const n = options?.headersRow ?? 1;
  return Math.max(0, n - 1);
}

function toStringHeaders(row: unknown[]): string[] {
  return row.map((c) => String(c ?? '').trim());
}

/**
 * Parse a 2D grid (Excel sheet values) into normalized invoice rows.
 */
export function parseInvoiceGrid(
  data: unknown[][],
  options: GridParseOptions = {},
): NormalizedInvoiceRow[] {
  if (!data?.length) return [];
  const hIdx = headerRowIndex(options);
  const headers = toStringHeaders(data[hIdx] ?? []);
  const mapping = mergeColumnMapping(headers, options.columnMapping);
  const documentId = options.documentId ?? 'workbook';
  const documentType = options.documentType ?? 'workbook';
  const skipEmpty = options.skipEmptyInvoice !== false;

  const invCol = resolveColumnIndex(mapping.invoiceNo, headers);
  const gstinCol = resolveColumnIndex(mapping.gstin, headers);
  if (invCol === undefined && gstinCol === undefined) {
    throw new Error(
      'Could not resolve invoice number or GSTIN columns. Provide column_mapping.',
    );
  }

  const rows: NormalizedInvoiceRow[] = [];
  for (let r = hIdx + 1; r < data.length; r++) {
    const row = data[r] ?? [];
    if (!Array.isArray(row)) continue;
    const allBlank = row.every(
      (c) => c === null || c === undefined || String(c).trim() === '',
    );
    if (allBlank) continue;

    const invoiceNumber = cellAt(row, headers, mapping, 'invoiceNo');
    const gstin = cellAt(row, headers, mapping, 'gstin');
    if (skipEmpty && !String(invoiceNumber ?? '').trim() && !String(gstin ?? '').trim()) {
      continue;
    }

    rows.push(
      buildNormalizedRow({
        gstin,
        invoiceNumber,
        invoiceDate: cellAt(row, headers, mapping, 'invoiceDate'),
        taxableValue: cellAt(row, headers, mapping, 'taxableAmt'),
        taxAmount: cellAt(row, headers, mapping, 'taxAmount'),
        igst: cellAt(row, headers, mapping, 'igst'),
        cgst: cellAt(row, headers, mapping, 'cgst'),
        sgst: cellAt(row, headers, mapping, 'sgst'),
        narration: cellAt(row, headers, mapping, 'narration'),
        documentType: cellAt(row, headers, mapping, 'documentType'),
        irn: cellAt(row, headers, mapping, 'irn'),
        imsAction: cellAt(row, headers, mapping, 'imsAction'),
        sourceRowRef: {
          documentType,
          documentId,
          rowOrLine: r + 1,
        },
      }),
    );
  }
  return rows;
}

export function parseRowsFromObjects(
  rows: Array<Record<string, unknown>>,
  documentType: DomainDocumentType,
  documentId: string,
): NormalizedInvoiceRow[] {
  return rows.map((raw, i) =>
    buildNormalizedRow({
      gstin: raw.gstin,
      invoiceNumber: raw.invoiceNumber ?? raw.invoiceNo,
      invoiceDate: raw.invoiceDate,
      taxableValue: raw.taxableValue ?? raw.taxableAmt,
      taxAmount: raw.taxAmount,
      igst: raw.igst,
      cgst: raw.cgst,
      sgst: raw.sgst,
      narration: raw.narration,
      documentType: raw.documentType,
      irn: raw.irn,
      imsAction: raw.imsAction,
      sourceRowRef: {
        documentType,
        documentId,
        rowOrLine: (raw.sourceRow as number) ?? i + 2,
      },
    }),
  );
}
