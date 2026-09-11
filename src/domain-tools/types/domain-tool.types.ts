/**
 * Shared contracts for deterministic domain tools.
 * The LLM plans/judges; these tools compute — never call an LLM inside.
 */

export type DomainDocumentType =
  | 'gstr2b'
  | 'gstr2a'
  | 'gstr1'
  | 'ims'
  | 'purchase_register'
  | 'sales_register'
  | 'form26as'
  | 'tally'
  | 'bank_statement'
  | 'workbook';

export interface SourceRef {
  documentType: DomainDocumentType;
  documentId: string;
  rowOrLine: string | number;
}

export interface DomainException {
  /** Versioned code, e.g. 'GST_NAME_FUZZY_MATCH' */
  code: string;
  /** 'block' prevents write; 'flag' allows write + review marker */
  severity: 'flag' | 'block';
  message: string;
  affectedRows: number[];
}

/**
 * confidence and exceptions are required — a tool cannot omit them.
 */
export interface DomainToolResult<T> {
  data: T;
  confidence: number;
  exceptions: DomainException[];
  sourceRefs: SourceRef[];
}

/** Deterministic tool signature — no LLM calls in the call graph. */
export type DomainTool<TInput, TOutput> = (input: TInput) => DomainToolResult<TOutput>;

export type GstDocumentType =
  | 'invoice'
  | 'credit_note'
  | 'debit_note'
  | 'amended'
  | 'unknown';

export type ImsActionStatus =
  | 'Accepted'
  | 'Rejected'
  | 'Pending'
  | 'AutoAccepted'
  | null;

export type SupplyCategory =
  | 'b2b'
  | 'b2c'
  | 'export'
  | 'sez'
  | 'nil_rated'
  | 'exempt';

/** Common normalized invoice row used by GST matching + portal ingestion. */
export interface NormalizedInvoiceRow {
  /** Counterparty GSTIN: supplier (purchase) or recipient (sales B2B). */
  gstin: string;
  invoiceNumber: string;
  /** Uppercase alphanumeric-only form used for exact key matches */
  normalizedInvoiceNumber: string;
  invoiceDate: string;
  /** null when derived from a rate-slab layout and no slab column had a value for this row. */
  taxableValue: number | null;
  /** Rate slab the taxable value was derived from, when the books sheet uses per-rate columns instead of a single Taxable Value column. */
  taxRatePercent?: number | null;
  /** True when 2+ rate-slab columns were non-blank and none's implied tax matched this row's actual tax closely enough to auto-select — taxableValue is null; needs CA review, never guessed. */
  ambiguousRateSlab?: boolean;
  ambiguousRateSlabDetail?: string;
  taxAmount: number;
  igst: number;
  cgst: number;
  sgst: number;
  narration: string;
  documentType: GstDocumentType;
  irn?: string;
  imsAction?: ImsActionStatus;
  /** Client-side GSTIN when present on the books register. */
  clientSideGstin?: string;
  supplyCategory?: SupplyCategory;
  placeOfSupply?: string;
  sourceRowRef: SourceRef;
}

export interface MatchedPair {
  registerRow: NormalizedInvoiceRow;
  portalRow: NormalizedInvoiceRow;
  matchKeys: string[];
  confidence: number;
  pass: number;
  difference?: string;
  diffType?: string;
}

export type GstReconStatus =
  | 'MATCHED'
  | 'PARTIAL'
  | 'CREDIT_NOTE'
  | 'PR_ONLY'
  | 'PORTAL_ONLY'
  | 'GSTIN_MISMATCH'
  | 'IMS_REJECTED'
  | 'IMS_PENDING'
  | 'IMS_AUTO_ACCEPT'
  | 'IMS_ONLY'
  | 'RCM'
  | 'AI_REVIEW';

/**
 * Specific reason a books row (PR_ONLY) didn't cleanly match the portal —
 * replaces a flat "unmatched" bucket so every row has a CA-readable cause.
 */
export type MismatchReason =
  | 'blank_counterparty_gstin'
  | 'blank_gstin_likely_matched'
  | 'blank_taxable_value'
  | 'ambiguous_rate_slab'
  | 'gstin_mismatch_same_pan'
  | 'gstin_not_in_portal'
  | 'amount_mismatch'
  | 'date_mismatch'
  | 'genuinely_missing';

export interface GstMatchSettings {
  amountToleranceAbs: number;
  amountTolerancePct: number;
  invoiceFuzzyThreshold: number;
  dateToleranceDays: number;
  detectRcm: boolean;
  useImsData: boolean;
}

export const DEFAULT_GST_MATCH_SETTINGS: GstMatchSettings = {
  amountToleranceAbs: 1.0,
  amountTolerancePct: 0.5,
  invoiceFuzzyThreshold: 85,
  dateToleranceDays: 3,
  detectRcm: true,
  useImsData: false,
};

export type PortalFileType =
  | 'GSTR2B'
  | 'GSTR2A'
  | 'GSTR1'
  | 'IMS'
  | 'PURCHASE_REGISTER'
  | 'SALES_REGISTER'
  | 'UNKNOWN';

export type ColumnMapping = Partial<{
  gstin: string | number;
  clientGstin: string | number;
  invoiceNo: string | number;
  invoiceDate: string | number;
  taxableAmt: string | number;
  taxAmount: string | number;
  igst: string | number;
  cgst: string | number;
  sgst: string | number;
  narration: string | number;
  irn: string | number;
  documentType: string | number;
  imsAction: string | number;
  supplyCategory: string | number;
  placeOfSupply: string | number;
}>;

/** Matching mode: purchase requires counterparty GSTIN; sales allows B2C blank. */
export type GstMatchMode = 'purchase' | 'sales';
