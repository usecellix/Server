/**
 * Shared contracts for deterministic domain tools.
 * The LLM plans/judges; these tools compute — never call an LLM inside.
 */

export type DomainDocumentType =
  | 'gstr2b'
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

/** Common normalized invoice row used by GST matching + GSTR-2B ingestion. */
export interface NormalizedInvoiceRow {
  gstin: string;
  invoiceNumber: string;
  invoiceDate: string;
  taxableValue: number;
  taxAmount: number;
  sourceRowRef: SourceRef;
}

export interface MatchedPair {
  registerRow: NormalizedInvoiceRow;
  portalRow: NormalizedInvoiceRow;
  matchKeys: string[];
  confidence: number;
}
