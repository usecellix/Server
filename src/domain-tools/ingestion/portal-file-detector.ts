import { ColumnMapping, PortalFileType } from '../types/domain-tool.types';

function normHeader(h: unknown): string {
  return String(h ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-./]+/g, ' ');
}

const HINTS: Record<Exclude<PortalFileType, 'UNKNOWN'>, string[]> = {
  GSTR2B: ['gstr 2b', 'gstr2b', '2b', 'itc available', 'autocrafted', 'b2b'],
  GSTR2A: ['gstr 2a', 'gstr2a', '2a', 'tds', 'tcs'],
  GSTR1: ['gstr 1', 'gstr1', 'gstr-1', 'b2b', 'b2cl', 'b2cs', 'outward'],
  IMS: ['ims action', 'ims status', 'invoice management', 'accept reject', 'ims'],
  PURCHASE_REGISTER: [
    'purchase register',
    'voucher',
    'party name',
    'supplier name',
    'bill no',
    'purchase',
  ],
  SALES_REGISTER: [
    'sales register',
    'sales',
    'outward',
    'recipient gstin',
    'customer name',
  ],
};

/**
 * Classify a grid (header row) as PR / GSTR-2B / GSTR-2A / IMS.
 */
export function detectPortalFileType(
  headers: unknown[],
  sheetName?: string,
): PortalFileType {
  const joined = [...headers.map(normHeader), normHeader(sheetName ?? '')].join(' | ');
  const name = normHeader(sheetName ?? '');

  if (name.includes('ims') || joined.includes('ims action') || joined.includes('ims status')) {
    return 'IMS';
  }
  if (name.includes('2a') || joined.includes('gstr 2a') || joined.includes('gstr2a')) {
    return 'GSTR2A';
  }
  if (
    name.includes('2b') ||
    joined.includes('gstr 2b') ||
    joined.includes('gstr2b') ||
    joined.includes('itc available')
  ) {
    return 'GSTR2B';
  }
  if (
    name.includes('gstr1') ||
    name.includes('gstr 1') ||
    /gstr-?1/.test(name) ||
    joined.includes('gstin of recipient') ||
    joined.includes('gstin/uin of recipient')
  ) {
    return 'GSTR1';
  }
  if (
    name.includes('sales') ||
    name.includes('outward') ||
    joined.includes('recipient gstin') ||
    joined.includes('customer gstin')
  ) {
    return 'SALES_REGISTER';
  }
  if (
    name.includes('purchase') ||
    joined.includes('party name') ||
    joined.includes('supplier name') ||
    joined.includes('voucher')
  ) {
    return 'PURCHASE_REGISTER';
  }

  let best: PortalFileType = 'UNKNOWN';
  let bestScore = 0;
  for (const [type, hints] of Object.entries(HINTS) as Array<
    [Exclude<PortalFileType, 'UNKNOWN'>, string[]]
  >) {
    let score = 0;
    for (const h of hints) {
      if (joined.includes(h)) score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      best = type;
    }
  }
  return bestScore > 0 ? best : 'UNKNOWN';
}

type LogicalCol = keyof ColumnMapping;

const HEADER_ALIASES: Record<LogicalCol, string[]> = {
  gstin: [
    'gstin of supplier',
    'supplier gstin',
    'gstin of recipient',
    'recipient gstin',
    'receiver gstin',
    'customer gstin',
    'gstin/uin of recipient',
    'gstin uin',
    'gstin/uin',
    'gstin',
    'gst no',
  ],
  clientGstin: [
    'client gstin',
    'our gstin',
    'company gstin',
    'entity gstin',
    'gstin of registered person',
  ],
  invoiceNo: [
    'invoice number',
    'invoice no',
    'invoice#',
    'inv no',
    'bill no',
    'voucher no',
    'voucher number',
    'doc no',
    'document number',
  ],
  invoiceDate: [
    'invoice date',
    'date',
    'bill date',
    'voucher date',
    'document date',
    'inv date',
  ],
  taxableAmt: [
    'taxable value',
    'taxable amount',
    'taxable amt',
    'taxable',
    'assessable value',
    'net amount',
  ],
  taxAmount: ['tax amount', 'total tax', 'integrated tax', 'gst amount', 'total gst'],
  igst: ['igst', 'igst amount', 'integrated tax amount'],
  cgst: ['cgst', 'cgst amount', 'central tax'],
  sgst: ['sgst', 'sgst amount', 'utgst', 'state tax'],
  narration: ['narration', 'description', 'particulars', 'remarks', 'item description', 'name'],
  irn: ['irn', 'invoice reference number', 'e invoice irn'],
  documentType: [
    'document type',
    'doc type',
    'invoice type',
    'voucher type',
    'type',
  ],
  imsAction: [
    'ims action',
    'ims status',
    'action status',
    'recipient action',
    'action',
    'status',
  ],
  supplyCategory: ['supply category', 'supply type', 'b2b/b2c', 'b2b', 'b2c'],
  placeOfSupply: ['place of supply', 'pos', 'state code'],
};

function colLetterToIndex(letter: string): number {
  const s = letter.trim().toUpperCase();
  if (/^\d+$/.test(s)) return Number(s);
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 65 || c > 90) return -1;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

/** Resolve Excel column letter or 0-based index to 0-based index. */
export function resolveColumnIndex(
  mapping: string | number | undefined,
  headers: string[],
): number | undefined {
  if (mapping === undefined || mapping === null || mapping === '') return undefined;
  if (typeof mapping === 'number') {
    if (mapping >= 0 && mapping < headers.length) return mapping;
    // 1-based column number
    if (mapping >= 1 && mapping <= headers.length) return mapping - 1;
    return undefined;
  }
  const asIndex = colLetterToIndex(mapping);
  if (asIndex >= 0 && asIndex < headers.length) return asIndex;
  const target = normHeader(mapping);
  const found = headers.findIndex((h) => normHeader(h) === target);
  return found >= 0 ? found : undefined;
}

/** True when the header row has a recognizable Invoice Number column. */
export function hasInvoiceNumberColumn(headers: unknown[]): boolean {
  const norms = headers.map((h) => normHeader(h));
  return HEADER_ALIASES.invoiceNo.some((alias) => norms.some((h) => h === alias || h.includes(alias)));
}

/**
 * Auto-detect column mapping from header labels.
 */
export function inferColumnMapping(headers: unknown[]): ColumnMapping {
  const norms = headers.map((h) => normHeader(h));
  const mapping: ColumnMapping = {};
  for (const [logical, aliases] of Object.entries(HEADER_ALIASES) as Array<
    [LogicalCol, string[]]
  >) {
    let idx = -1;
    for (const alias of aliases) {
      idx = norms.findIndex((h) => h === alias || h.includes(alias));
      if (idx >= 0) break;
    }
    if (idx >= 0) {
      mapping[logical] = idx;
    }
  }
  return mapping;
}

export function mergeColumnMapping(
  headers: unknown[],
  explicit?: ColumnMapping | null,
): ColumnMapping {
  const auto = inferColumnMapping(headers);
  if (!explicit) return auto;
  return { ...auto, ...explicit };
}

export function cellAt(
  row: unknown[],
  headers: string[],
  mapping: ColumnMapping,
  key: LogicalCol,
): unknown {
  const idx = resolveColumnIndex(mapping[key], headers);
  if (idx === undefined) return undefined;
  return row[idx];
}
