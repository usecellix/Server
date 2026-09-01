/**
 * Shared contracts for deterministic domain tools.
 * The LLM plans/judges; these tools compute — never call an LLM inside.
 */

export type DomainDocumentType =
  | 'gstr2b'
  | 'gstr2a'
  | 'ims'
  | 'purchase_register'
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

/** Common normalized invoice row used by GST matching + portal ingestion. */
export interface NormalizedInvoiceRow {
  gstin: string;
  invoiceNumber: string;
  /** Uppercase alphanumeric-only form used for exact key matches */
  normalizedInvoiceNumber: string;
  invoiceDate: string;
  taxableValue: number;
  taxAmount: number;
  igst: number;
  cgst: number;
  sgst: number;
  narration: string;
  documentType: GstDocumentType;
  irn?: string;
  imsAction?: ImsActionStatus;
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
  | 'IMS_REJECTED'
  | 'IMS_PENDING'
  | 'IMS_AUTO_ACCEPT'
  | 'IMS_ONLY'
  | 'RCM'
  | 'AI_REVIEW';

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

export type PortalFileType = 'GSTR2B' | 'GSTR2A' | 'IMS' | 'PURCHASE_REGISTER' | 'UNKNOWN';

export type ColumnMapping = Partial<{
  gstin: string | number;
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
}>;
