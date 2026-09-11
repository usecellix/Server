/**
 * Shared contracts for prompt-based Purchase & Sales GST reconciliation.
 * clientGstin is the single required run-level key for both recon types.
 */

export type ReconType = 'purchase_vs_2b' | 'purchase_vs_2a' | 'sales_vs_gstr1';

export type GstRole = 'recipient' | 'issuer';

export const RECON_TYPE_ROLE: Record<ReconType, GstRole> = {
  purchase_vs_2b: 'recipient',
  purchase_vs_2a: 'recipient',
  sales_vs_gstr1: 'issuer',
};

/** Map API DTO recon types to internal ReconType. */
export function dtoToReconType(dto: string): ReconType | null {
  switch (dto) {
    case 'PR_VS_GSTR2B':
      return 'purchase_vs_2b';
    case 'PR_VS_GSTR2A':
      return 'purchase_vs_2a';
    case 'SALES_VS_GSTR1':
      return 'sales_vs_gstr1';
    default:
      return null;
  }
}

export function reconTypeToDto(t: ReconType): string {
  switch (t) {
    case 'purchase_vs_2b':
      return 'PR_VS_GSTR2B';
    case 'purchase_vs_2a':
      return 'PR_VS_GSTR2A';
    case 'sales_vs_gstr1':
      return 'SALES_VS_GSTR1';
  }
}

export interface ReconRunContext {
  operator?: { id: string; name: string };
  firm?: { id: string; name: string };
  client?: { id: string; name: string };
  clientGstin: string;
  reconType: ReconType;
  financialYear?: string;
  taxPeriod: string;
  sourceExtractionDate?: string;
}

export interface SheetCandidate {
  sheetName: string;
  headers: string[];
  headerRowIndex: number;
  confidence: number;
  detectedGstin?: string;
}

export type RegisterSignature =
  | 'purchase_register'
  | 'sales_register'
  | 'gstr_2a'
  | 'gstr_2b'
  | 'gstr_1';

export type SupplyCategory =
  | 'b2b'
  | 'b2c'
  | 'export'
  | 'sez'
  | 'nil_rated'
  | 'exempt';

export interface NormalizedRow {
  rowIndex: number;
  /** Supplier GSTIN (purchase) or recipient GSTIN (sales B2B). */
  counterpartyGstin: string | null;
  /** Client-side GSTIN column when present (recipient in PR, issuer in SR). */
  clientSideGstin?: string | null;
  documentType: 'invoice' | 'credit_note' | 'debit_note' | 'amended_invoice';
  invoiceNumber: string;
  invoiceNumberRaw: string;
  invoiceDate: string;
  taxableValue: number;
  igst: number;
  cgst: number;
  sgst: number;
  cess: number;
  totalValue?: number;
  placeOfSupply?: string;
  supplyCategory?: SupplyCategory;
  irn?: string;
  ewayBillNo?: string;
  itcEligibility?: string;
}

export type MatchStatus =
  | 'matched'
  | 'partial_match'
  | 'credit_debit_note_match'
  | 'books_only'
  | 'portal_only'
  | 'duplicate_suspected'
  | 'invalid_missing_gstin'
  | 'cross_gstin_exception'
  | 'period_mismatch'
  | 'llm_review';

export interface MatchResult {
  status: MatchStatus;
  pass: 1 | 2 | 3 | 6 | null;
  booksRow?: NormalizedRow;
  portalRow?: NormalizedRow;
  differences?: Record<string, { books: unknown; portal: unknown }>;
  confidence?: number;
  reviewReason?: string;
}

export interface ReconSummary {
  reconType: ReconType;
  clientGstin: string;
  period: string;
  counts: Partial<Record<MatchStatus, number>>;
  taxableValueMatched: number;
  taxAmountMatched: number;
  taxAmountAtRisk: number;
  crossGstinExceptionCount: number;
  generatedAt: string;
}

export interface GstinValidationResult {
  ok: boolean;
  normalizedGstin?: string;
  blockingError?: string;
  warning?: string;
}

export type NormalizedRowWithClientGstinColumn = NormalizedRow & {
  clientSideGstin?: string | null;
};

export interface ReconActionPayload {
  actionType: 'GST_RECON_RESULT';
  reconType: ReconType;
  clientGstin: string;
  clientName: string;
  period: string;
  summary: ReconSummary;
  crossGstinExceptionCount: number;
  suggestedSheetName: string;
  writeAction: {
    type: 'CREATE_SHEET_AND_WRITE_TABLE';
    sheetName: string;
    sections: Array<'header_block' | 'summary' | 'detail_rows'>;
  };
}

export interface GstReconAuditEntry {
  operatorId?: string;
  firmId?: string;
  clientId?: string;
  clientGstin: string;
  reconType: ReconType;
  period: string;
  sourceSheets: { books: string; portal: string };
  sourceFileHashes?: { books: string; portal: string };
  matchingSettings?: Record<string, unknown>;
  crossGstinExceptionCount: number;
  outcome: 'pending' | 'applied' | 'rejected';
  appliedAt?: string;
}
