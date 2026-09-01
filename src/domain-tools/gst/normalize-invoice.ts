import {
  GstDocumentType,
  ImsActionStatus,
  NormalizedInvoiceRow,
  SourceRef,
} from '../types/domain-tool.types';

/** Strip separators and lowercase for fuzzy compare; keep alphanumerics. */
export function normalizeInvoiceNumber(raw: string): string {
  return String(raw ?? '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

export function normalizeGstin(raw: string): string {
  return String(raw ?? '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '');
}

export function parseAmount(raw: unknown): number {
  if (raw === null || raw === undefined || raw === '') return 0;
  if (typeof raw === 'number' && Number.isFinite(raw)) return roundMoney(raw);
  const cleaned = String(raw)
    .replace(/₹/g, '')
    .replace(/,/g, '')
    .replace(/\s+/g, '')
    .trim();
  const n = Number(cleaned);
  return Number.isFinite(n) ? roundMoney(n) : 0;
}

export function roundMoney(n: number): number {
  return Math.round(n * 100) / 100;
}

/** Parse common Indian date formats to ISO-ish yyyy-mm-dd when possible. */
export function normalizeDate(raw: unknown): string {
  if (raw === null || raw === undefined) return '';
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return raw.toISOString().slice(0, 10);
  }
  const s = String(raw).trim();
  if (!s) return '';
  // Excel serial date (approx)
  if (/^\d+(\.\d+)?$/.test(s)) {
    const serial = Number(s);
    if (serial > 20000 && serial < 60000) {
      const utc = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
      return utc.toISOString().slice(0, 10);
    }
  }
  const dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (dmy) {
    const day = dmy[1].padStart(2, '0');
    const month = dmy[2].padStart(2, '0');
    let year = dmy[3];
    if (year.length === 2) year = `20${year}`;
    return `${year}-${month}-${day}`;
  }
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  return s;
}

export function amountWithinTolerance(
  a: number,
  b: number,
  absTol: number,
  pctTol: number,
): boolean {
  const diff = Math.abs(a - b);
  if (diff <= absTol + 1e-9) return true;
  const base = Math.max(Math.abs(a), Math.abs(b), 1e-9);
  return (diff / base) * 100 <= pctTol + 1e-9;
}

export function daysBetween(dateA: string, dateB: string): number | null {
  if (!dateA || !dateB) return null;
  const ta = Date.parse(dateA);
  const tb = Date.parse(dateB);
  if (Number.isNaN(ta) || Number.isNaN(tb)) return null;
  return Math.abs(Math.round((ta - tb) / 86400000));
}

export function detectDocumentType(
  rawType: unknown,
  narration: string,
  taxableValue: number,
): GstDocumentType {
  const t = String(rawType ?? '')
    .trim()
    .toUpperCase();
  if (t === 'C' || t === 'CR' || t === 'CN' || t.includes('CREDIT') || t === 'CDN') {
    return 'credit_note';
  }
  if (t === 'D' || t === 'DR' || t === 'DN' || t.includes('DEBIT')) {
    return 'debit_note';
  }
  if (t === 'IA' || t.includes('AMEND')) {
    return 'amended';
  }
  if (t === 'I' || t === 'INV' || t.includes('INVOICE') || t === 'B2B') {
    return 'invoice';
  }
  const n = narration.toUpperCase();
  if (n.includes('CREDIT NOTE') || n.includes('CREDITNOTE') || /\bCDN\b/.test(n)) {
    return 'credit_note';
  }
  if (n.includes('DEBIT NOTE') || n.includes('DEBITNOTE')) {
    return 'debit_note';
  }
  if (taxableValue < 0 || parseAmount(taxableValue) < 0) {
    return 'credit_note';
  }
  return t ? 'unknown' : 'invoice';
}

export function parseImsAction(raw: unknown): ImsActionStatus {
  if (raw === null || raw === undefined || raw === '') return null;
  const s = String(raw).trim().toLowerCase();
  if (!s) return null;
  if (s.includes('reject')) return 'Rejected';
  if (s.includes('pending')) return 'Pending';
  if (s.includes('auto')) return 'AutoAccepted';
  if (s.includes('accept') || s === 'a' || s === 'y' || s === 'yes') return 'Accepted';
  if (s.includes('no action') || s.includes('nil')) return 'AutoAccepted';
  return null;
}

export function isValidGstinFormat(gstin: string): boolean {
  if (!gstin) return false;
  return /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][A-Z0-9]Z[A-Z0-9]$/.test(gstin);
}

export function buildNormalizedRow(parts: {
  gstin: unknown;
  invoiceNumber: unknown;
  invoiceDate: unknown;
  taxableValue: unknown;
  taxAmount?: unknown;
  igst?: unknown;
  cgst?: unknown;
  sgst?: unknown;
  narration?: unknown;
  documentType?: unknown;
  irn?: unknown;
  imsAction?: unknown;
  sourceRowRef: SourceRef;
}): NormalizedInvoiceRow {
  const gstin = normalizeGstin(String(parts.gstin ?? ''));
  const invoiceNumber = String(parts.invoiceNumber ?? '').trim();
  const taxableValue = parseAmount(parts.taxableValue);
  const igst = parseAmount(parts.igst);
  const cgst = parseAmount(parts.cgst);
  const sgst = parseAmount(parts.sgst);
  let taxAmount = parseAmount(parts.taxAmount);
  if (!taxAmount && (igst || cgst || sgst)) {
    taxAmount = roundMoney(igst + cgst + sgst);
  }
  const narration = String(parts.narration ?? '').trim();
  const documentType = detectDocumentType(parts.documentType, narration, taxableValue);
  const irn = parts.irn != null && String(parts.irn).trim() ? String(parts.irn).trim() : undefined;
  const imsAction =
    parts.imsAction !== undefined ? parseImsAction(parts.imsAction) : undefined;

  return {
    gstin,
    invoiceNumber,
    normalizedInvoiceNumber: normalizeInvoiceNumber(invoiceNumber),
    invoiceDate: normalizeDate(parts.invoiceDate),
    taxableValue,
    taxAmount,
    igst,
    cgst,
    sgst,
    narration,
    documentType,
    irn,
    imsAction: imsAction === undefined ? undefined : imsAction,
    sourceRowRef: parts.sourceRowRef,
  };
}

/** Simple Levenshtein distance for fuzzy invoice matching. */
export function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const prev = new Array<number>(b.length + 1);
  const curr = new Array<number>(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    for (let j = 0; j <= b.length; j++) prev[j] = curr[j];
  }
  return prev[b.length];
}

/** Similarity 0–100 based on Levenshtein. */
export function stringSimilarityPercent(a: string, b: string): number {
  if (!a && !b) return 100;
  if (!a || !b) return 0;
  if (a === b) return 100;
  const dist = levenshtein(a, b);
  const maxLen = Math.max(a.length, b.length);
  return Math.round((1 - dist / maxLen) * 10000) / 100;
}
