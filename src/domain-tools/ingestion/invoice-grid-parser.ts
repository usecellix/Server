import {
  ColumnMapping,
  DomainDocumentType,
  NormalizedInvoiceRow,
} from '../types/domain-tool.types';
import {
  buildNormalizedRow,
  deriveTaxableValueFromSlabRow,
  isRateSlabLayout,
  parseAmount,
  roundMoney,
} from '../gst/normalize-invoice';
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

  const slabLayout = isRateSlabLayout(headers);

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
    // Only treat "blank invoice + blank GSTIN" as a junk/subtotal row when the sheet actually
    // has a resolvable Invoice Number column — otherwise every real row has a blank invoice
    // number by definition, and this would silently drop genuine blank-GSTIN rows before they
    // ever reach mismatch diagnosis (the exact "silently falls through" failure this pipeline
    // exists to prevent).
    if (
      skipEmpty &&
      invCol !== undefined &&
      !String(invoiceNumber ?? '').trim() &&
      !String(gstin ?? '').trim()
    ) {
      continue;
    }

    let slabTaxableValue: number | null = null;
    let slabTaxRatePercent: number | null = null;
    let slabAmbiguous = false;
    let slabAmbiguousDetail: string | undefined;
    if (slabLayout) {
      const rawIgst = parseAmount(cellAt(row, headers, mapping, 'igst'));
      const rawCgst = parseAmount(cellAt(row, headers, mapping, 'cgst'));
      const rawSgst = parseAmount(cellAt(row, headers, mapping, 'sgst'));
      const actualTax = rawIgst || roundMoney(rawCgst + rawSgst);
      const derived = deriveTaxableValueFromSlabRow(
        Object.fromEntries(headers.map((h, i) => [h, row[i]])),
        actualTax,
      );
      if (derived && 'ambiguous' in derived) {
        slabAmbiguous = true;
        slabAmbiguousDetail = derived.detail;
      } else if (derived) {
        slabTaxableValue = derived.taxableValue;
        slabTaxRatePercent = derived.taxRatePercent;
      }
    }

    rows.push(
      buildNormalizedRow({
        gstin,
        invoiceNumber,
        invoiceDate: cellAt(row, headers, mapping, 'invoiceDate'),
        taxableValue: slabLayout ? slabTaxableValue : cellAt(row, headers, mapping, 'taxableAmt'),
        taxRatePercent: slabLayout ? slabTaxRatePercent : undefined,
        ambiguousRateSlab: slabAmbiguous ? true : undefined,
        ambiguousRateSlabDetail: slabAmbiguousDetail,
        taxAmount: cellAt(row, headers, mapping, 'taxAmount'),
        igst: cellAt(row, headers, mapping, 'igst'),
        cgst: cellAt(row, headers, mapping, 'cgst'),
        sgst: cellAt(row, headers, mapping, 'sgst'),
        narration: cellAt(row, headers, mapping, 'narration'),
        documentType: cellAt(row, headers, mapping, 'documentType'),
        irn: cellAt(row, headers, mapping, 'irn'),
        imsAction: cellAt(row, headers, mapping, 'imsAction'),
        clientSideGstin: cellAt(row, headers, mapping, 'clientGstin'),
        supplyCategory: cellAt(row, headers, mapping, 'supplyCategory'),
        placeOfSupply: cellAt(row, headers, mapping, 'placeOfSupply'),
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
