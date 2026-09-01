import {
  DomainTool,
  DomainToolResult,
  NormalizedInvoiceRow,
  SourceRef,
} from '../types/domain-tool.types';
import { GstResultRow } from './match-passes';
import { roundMoney } from './normalize-invoice';

export interface ItcComputeInput {
  /** Prefer full recon result rows when available */
  resultRows?: GstResultRow[];
  /** Fallback: invoices treated as eligible ITC */
  eligibleInvoices?: NormalizedInvoiceRow[];
  /** e.g. 0.18 — reserved for rate reconstruction; not required when tax amounts present */
  igstRate?: number;
  cgstRate?: number;
  sgstRate?: number;
}

export interface ItcLineResult {
  invoiceNumber: string;
  taxableValue: number;
  itcClaimable: number;
  status: string;
  sourceRowRef: SourceRef;
}

export interface ItcComputeOutput {
  lines: ItcLineResult[];
  totalItcClaimable: number;
  totalItcAtRisk: number;
  totalRcmPayable: number;
  totalImsRejected: number;
  totalImsPending: number;
  totalMatched: number;
  totalPartial: number;
  totalPrOnly: number;
  totalPortalOnly: number;
  totalCreditNotes: number;
  totalRcm: number;
}

function taxOf(row: NormalizedInvoiceRow): number {
  return Math.abs(row.taxAmount || roundMoney(row.igst + row.cgst + row.sgst));
}

/**
 * ITC summary from recon outcomes — pure arithmetic, never LLM.
 */
export const itcCompute: DomainTool<ItcComputeInput, ItcComputeOutput> = (input) => {
  const lines: ItcLineResult[] = [];
  let totalItcClaimable = 0;
  let totalItcAtRisk = 0;
  let totalRcmPayable = 0;
  let totalImsRejected = 0;
  let totalImsPending = 0;
  let totalMatched = 0;
  let totalPartial = 0;
  let totalPrOnly = 0;
  let totalPortalOnly = 0;
  let totalCreditNotes = 0;
  let totalRcm = 0;
  const sourceRefs: SourceRef[] = [];

  if (input.resultRows?.length) {
    for (const r of input.resultRows) {
      const row = r.registerRow ?? r.portalRow ?? r.imsRow;
      if (!row) continue;
      sourceRefs.push(row.sourceRowRef);

      let itcClaimable = 0;
      if (r.status === 'MATCHED') {
        itcClaimable = taxOf(row);
        totalItcClaimable += itcClaimable;
        totalMatched += 1;
      } else if (r.status === 'PARTIAL' || r.status === 'IMS_AUTO_ACCEPT') {
        itcClaimable = taxOf(row);
        // Partial still counted as claimable but flagged via match exceptions
        totalItcClaimable += itcClaimable;
        totalPartial += 1;
      } else if (r.status === 'CREDIT_NOTE') {
        itcClaimable = -taxOf(row);
        totalItcClaimable += itcClaimable;
        totalCreditNotes += 1;
      } else if (r.status === 'PR_ONLY') {
        totalItcAtRisk += taxOf(row);
        totalPrOnly += 1;
      } else if (r.status === 'PORTAL_ONLY' || r.status === 'IMS_ONLY') {
        totalPortalOnly += 1;
      } else if (r.status === 'IMS_REJECTED') {
        totalImsRejected += taxOf(row);
        totalItcAtRisk += taxOf(row);
      } else if (r.status === 'IMS_PENDING') {
        totalImsPending += taxOf(row);
      } else if (r.status === 'RCM') {
        // Approximate RCM payable as 18% of taxable when tax not present
        const rcm =
          taxOf(row) ||
          roundMoney(Math.abs(row.taxableValue) * (input.igstRate ?? input.cgstRate ?? 0.18));
        totalRcmPayable += rcm;
        totalRcm += 1;
      }

      lines.push({
        invoiceNumber: row.invoiceNumber,
        taxableValue: row.taxableValue,
        itcClaimable,
        status: r.status,
        sourceRowRef: row.sourceRowRef,
      });
    }
  } else {
    for (const inv of input.eligibleInvoices ?? []) {
      const itc = taxOf(inv);
      totalItcClaimable += itc;
      totalMatched += 1;
      sourceRefs.push(inv.sourceRowRef);
      lines.push({
        invoiceNumber: inv.invoiceNumber,
        taxableValue: inv.taxableValue,
        itcClaimable: itc,
        status: 'MATCHED',
        sourceRowRef: inv.sourceRowRef,
      });
    }
  }

  const result: DomainToolResult<ItcComputeOutput> = {
    data: {
      lines,
      totalItcClaimable: roundMoney(totalItcClaimable),
      totalItcAtRisk: roundMoney(totalItcAtRisk),
      totalRcmPayable: roundMoney(totalRcmPayable),
      totalImsRejected: roundMoney(totalImsRejected),
      totalImsPending: roundMoney(totalImsPending),
      totalMatched,
      totalPartial,
      totalPrOnly,
      totalPortalOnly,
      totalCreditNotes,
      totalRcm,
    },
    confidence: 1,
    exceptions: [],
    sourceRefs,
  };
  return result;
};
