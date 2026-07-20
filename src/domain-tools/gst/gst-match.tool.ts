import {
  DomainTool,
  MatchedPair,
  NormalizedInvoiceRow,
} from '../types/domain-tool.types';

export interface GstMatchInput {
  purchaseRegister: NormalizedInvoiceRow[];
  gstr2b: NormalizedInvoiceRow[];
  matchKeys: Array<'gstin' | 'invoiceNumber' | 'invoiceDate'>;
  amountTolerance: number;
}

export interface GstMatchOutput {
  matched: MatchedPair[];
  partialMatch: MatchedPair[];
  missingIn2B: NormalizedInvoiceRow[];
  missingInRegister: NormalizedInvoiceRow[];
}

/**
 * STUB — exact-key match first, then fuzzy vendor-name fallback with 'flag'
 * exceptions only. Requires CA-reviewed matching spec before production use.
 */
export const gstMatch: DomainTool<GstMatchInput, GstMatchOutput> = (_input) => {
  throw new Error(
    'Not implemented — requires CA-reviewed matching spec before production use.',
  );
};
