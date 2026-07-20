import { DomainTool, NormalizedInvoiceRow, SourceRef } from '../types/domain-tool.types';

export interface ItcComputeInput {
  eligibleInvoices: NormalizedInvoiceRow[];
  /** e.g. 0.18 for 18% — rates must come from caller config, not LLM free text */
  igstRate?: number;
  cgstRate?: number;
  sgstRate?: number;
}

export interface ItcLineResult {
  invoiceNumber: string;
  taxableValue: number;
  itcClaimable: number;
  sourceRowRef: SourceRef;
}

export interface ItcComputeOutput {
  lines: ItcLineResult[];
  totalItcClaimable: number;
}

/**
 * STUB — ITC arithmetic must be CA-signed off; never LLM-computed at runtime.
 */
export const itcCompute: DomainTool<ItcComputeInput, ItcComputeOutput> = (_input) => {
  throw new Error(
    'Not implemented — requires CA-reviewed ITC computation spec before production use.',
  );
};
