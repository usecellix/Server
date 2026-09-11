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

/**
 * Books exports sometimes split taxable value across per-rate columns
 * (e.g. "Purchase@18%") instead of a single Taxable Value column.
 */
export const RATE_SLAB_COLUMNS = [
  'Purchase@0%',
  'Purchase@5%',
  'Purchase@12%',
  'Purchase@18%',
  'Purchase@28%',
  'Purchase Interstate@0%',
  'Purchase Interstate@5%',
  'Purchase Interstate@12%',
  'Purchase Interstate@18%',
  'Purchase Interstate@28%',
  'Sales@0%',
  'Sales@5%',
  'Sales@12%',
  'Sales@18%',
  'Sales@28%',
  'Sales Interstate@0%',
  'Sales Interstate@5%',
  'Sales Interstate@12%',
  'Sales Interstate@18%',
  'Sales Interstate@28%',
];

/** Lowercase and strip ALL whitespace so "Purchase @ 18%" / "Purchase@18%" / "Purchase  @18 %" all compare equal. */
function normalizeHeader(h: unknown): string {
  return String(h ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');
}

/** True when the header row has 2+ rate-slab columns (no single Taxable Value column). */
export function isRateSlabLayout(headers: string[]): boolean {
  const matches = headers.filter((h) =>
    RATE_SLAB_COLUMNS.some((c) => normalizeHeader(h) === normalizeHeader(c)),
  );
  return matches.length >= 2;
}

/** "Purchase Interstate@18%" -> 18. Returns null when no rate can be parsed. */
export function extractRateFromColumnName(column: string): number | null {
  const m = column.match(/@\s*(\d+(?:\.\d+)?)\s*%/);
  return m ? Number(m[1]) : null;
}

export interface SlabCandidate {
  sourceColumn: string;
  taxableValue: number;
  taxRatePercent: number | null;
  /** taxableValue * taxRatePercent / 100 — what this candidate implies the row's tax should be. */
  impliedTax: number | null;
}

export interface SlabDerivation {
  taxableValue: number;
  taxRatePercent: number | null;
  sourceColumn: string;
}

export interface AmbiguousSlabDerivation {
  ambiguous: true;
  candidates: SlabCandidate[];
  detail: string;
}

/**
 * A row's implied tax (candidate taxableValue × its rate) must land within this of the
 * row's actual CGST+SGST/IGST to be auto-selected — loose enough for ordinary paise-level
 * rounding noise, tight enough to reject a genuinely wrong slab (real cases differ by
 * 5%+ once the wrong column is picked).
 */
const RATE_SLAB_TAX_TOLERANCE_ABS = 2;
const RATE_SLAB_TAX_TOLERANCE_PCT = 1;

function formatCandidate(c: SlabCandidate): string {
  return `${c.sourceColumn}: ₹${c.taxableValue}${c.impliedTax != null ? ` (implies tax ₹${c.impliedTax})` : ''}`;
}

/**
 * Given a row keyed by header name, find the non-blank, non-zero rate-slab column(s) and
 * derive taxable value + tax rate. When exactly one column is populated, use it directly.
 * When more than one is populated (a real data-entry pattern: a stray value left in an
 * unused rate column alongside the real one), pick the candidate whose implied tax
 * (value × rate) is closest to the row's actual tax — never the first one in priority
 * order. If no candidate's implied tax is reasonably close, don't guess: report it as
 * ambiguous for CA review.
 *
 * Returns null when no slab column has a value at all — a real data gap, not a mapping failure.
 */
export function deriveTaxableValueFromSlabRow(
  row: Record<string, unknown>,
  actualTax?: number | null,
): SlabDerivation | AmbiguousSlabDerivation | null {
  const byNormalizedHeader = new Map<string, string>();
  for (const key of Object.keys(row)) {
    byNormalizedHeader.set(normalizeHeader(key), key);
  }

  const candidates: SlabCandidate[] = [];
  for (const col of RATE_SLAB_COLUMNS) {
    const actualKey = byNormalizedHeader.get(normalizeHeader(col));
    if (actualKey === undefined) continue;
    const v = row[actualKey];
    if (v === null || v === undefined || v === '') continue;
    const n = Number(v);
    if (!Number.isFinite(n) || n === 0) continue;
    const taxableValue = roundMoney(n);
    const rate = extractRateFromColumnName(col);
    candidates.push({
      sourceColumn: col,
      taxableValue,
      taxRatePercent: rate,
      impliedTax: rate != null ? roundMoney((taxableValue * rate) / 100) : null,
    });
  }

  if (candidates.length === 0) return null;
  if (candidates.length === 1) {
    const c = candidates[0];
    return { taxableValue: c.taxableValue, taxRatePercent: c.taxRatePercent, sourceColumn: c.sourceColumn };
  }

  const scored = candidates
    .map((c) => ({
      ...c,
      diff: actualTax != null && c.impliedTax != null ? Math.abs(c.impliedTax - actualTax) : Number.POSITIVE_INFINITY,
    }))
    .sort((a, b) => a.diff - b.diff);
  const best = scored[0];

  if (
    actualTax != null &&
    actualTax > 0 &&
    best.impliedTax != null &&
    amountWithinTolerance(best.impliedTax, actualTax, RATE_SLAB_TAX_TOLERANCE_ABS, RATE_SLAB_TAX_TOLERANCE_PCT)
  ) {
    return { taxableValue: best.taxableValue, taxRatePercent: best.taxRatePercent, sourceColumn: best.sourceColumn };
  }

  const detail =
    `Multiple rate-slab columns are populated (${candidates.map(formatCandidate).join('; ')}) ` +
    (actualTax != null && actualTax > 0
      ? `but none's implied tax matches this row's actual tax (₹${actualTax}) closely enough to pick automatically.`
      : `and the row's actual tax amount could not be used to disambiguate.`);
  return { ambiguous: true, candidates, detail };
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
  taxableValue: number | null,
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
  if (taxableValue !== null && (taxableValue < 0 || parseAmount(taxableValue) < 0)) {
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
  /** Pass literal `null` (not `undefined`/blank) to mean "no value found" — kept as null, never defaulted to 0. */
  taxableValue: unknown;
  taxRatePercent?: number | null;
  ambiguousRateSlab?: boolean;
  ambiguousRateSlabDetail?: string;
  taxAmount?: unknown;
  igst?: unknown;
  cgst?: unknown;
  sgst?: unknown;
  narration?: unknown;
  documentType?: unknown;
  irn?: unknown;
  imsAction?: unknown;
  clientSideGstin?: unknown;
  supplyCategory?: unknown;
  placeOfSupply?: unknown;
  sourceRowRef: SourceRef;
}): NormalizedInvoiceRow {
  const gstin = normalizeGstin(String(parts.gstin ?? ''));
  const invoiceNumber = String(parts.invoiceNumber ?? '').trim();
  const taxableValue = parts.taxableValue === null ? null : parseAmount(parts.taxableValue);
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
  const clientSideGstin = parts.clientSideGstin
    ? normalizeGstin(String(parts.clientSideGstin))
    : undefined;
  const supplyCategory = classifySupplyCategory(parts.supplyCategory, gstin);
  const placeOfSupply = parts.placeOfSupply
    ? String(parts.placeOfSupply).trim()
    : undefined;

  return {
    gstin,
    invoiceNumber,
    normalizedInvoiceNumber: normalizeInvoiceNumber(invoiceNumber),
    invoiceDate: normalizeDate(parts.invoiceDate),
    taxableValue,
    taxRatePercent: parts.taxRatePercent ?? undefined,
    ambiguousRateSlab: parts.ambiguousRateSlab ?? undefined,
    ambiguousRateSlabDetail: parts.ambiguousRateSlabDetail ?? undefined,
    taxAmount,
    igst,
    cgst,
    sgst,
    narration,
    documentType,
    irn,
    imsAction: imsAction === undefined ? undefined : imsAction,
    clientSideGstin: clientSideGstin || undefined,
    supplyCategory,
    placeOfSupply,
    sourceRowRef: parts.sourceRowRef,
  };
}

export function classifySupplyCategory(
  raw: unknown,
  counterpartyGstin: string,
): import('../types/domain-tool.types').SupplyCategory | undefined {
  const s = String(raw ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-]/g, '');
  if (s.includes('b2c') || s === 'b2cl' || s === 'b2cs') return 'b2c';
  if (s.includes('export')) return 'export';
  if (s.includes('sez')) return 'sez';
  if (s.includes('nil')) return 'nil_rated';
  if (s.includes('exempt')) return 'exempt';
  if (s.includes('b2b')) return 'b2b';
  // Fallback: blank counterparty GSTIN ⇒ B2C for sales contexts
  if (!counterpartyGstin) return 'b2c';
  if (raw == null || String(raw).trim() === '') return 'b2b';
  return undefined;
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
