import { BadRequestException, Injectable } from '@nestjs/common';
import { randomUUID } from 'crypto';
import {
  ColumnMapping,
  DEFAULT_GST_MATCH_SETTINGS,
  GstMatchSettings,
} from '../domain-tools/types/domain-tool.types';
import { parsePurchaseRegister } from '../domain-tools/ingestion/purchase-register-parser';
import { parseGstr2a, parseGstr2b } from '../domain-tools/ingestion/gstr2b-parser';
import { parseImsExport } from '../domain-tools/ingestion/ims-parser';
import { gstMatch } from '../domain-tools/gst/gst-match.tool';
import { itcCompute } from '../domain-tools/gst/itc-compute.tool';
import { gstr3bVs2b } from '../domain-tools/gst/gstr3b-vs-2b.tool';
import { GstReconcileRequestDto, SheetPayloadDto } from './gst-recon.dto';
import {
  buildSummary,
  map3bVs2bToActions,
  mapResultRows,
  mapToSheetActions,
  ReconSummaryBlock,
} from './gst-recon-to-actions.mapper';

function toColumnMapping(raw?: SheetPayloadDto['column_mapping']): ColumnMapping | undefined {
  if (!raw) return undefined;
  const m: ColumnMapping = {};
  const assign = (
    key: keyof ColumnMapping,
    ...aliases: Array<string | number | undefined | null>
  ) => {
    for (const a of aliases) {
      if (a !== undefined && a !== null && a !== '') {
        m[key] = a;
        return;
      }
    }
  };
  assign('gstin', raw.gstin);
  assign('invoiceNo', raw.invoiceNo, raw.invoice_no);
  assign('invoiceDate', raw.invoiceDate);
  assign('taxableAmt', raw.taxableAmt, raw.taxable_amt);
  assign('taxAmount', raw.taxAmount);
  assign('igst', raw.igst);
  assign('cgst', raw.cgst);
  assign('sgst', raw.sgst);
  assign('narration', raw.narration);
  assign('irn', raw.irn);
  assign('documentType', raw.documentType);
  assign('imsAction', raw.imsAction);
  return Object.keys(m).length ? m : undefined;
}

function settingsFromDto(dto?: GstReconcileRequestDto['settings']): Partial<GstMatchSettings> {
  if (!dto) return {};
  return {
    amountToleranceAbs: dto.amount_tolerance_abs,
    amountTolerancePct: dto.amount_tolerance_pct,
    invoiceFuzzyThreshold: dto.invoice_fuzzy_threshold,
    dateToleranceDays: dto.date_tolerance_days,
    detectRcm: dto.detect_rcm,
    useImsData: dto.use_ims_data,
  };
}

@Injectable()
export class GstReconService {
  reconcile(request: GstReconcileRequestDto) {
    const jobId = `recon_${randomUUID().replace(/-/g, '').slice(0, 12)}`;
    const runAt = new Date().toISOString();

    if (request.reconciliation_type === 'GSTR3B_VS_GSTR2B') {
      return this.reconcile3bVs2b(request, jobId, runAt);
    }

    return this.reconcileInvoiceLevel(request, jobId, runAt);
  }

  private reconcile3bVs2b(request: GstReconcileRequestDto, jobId: string, runAt: string) {
    if (!request.gstr3b_itc || !request.gstr2b_itc) {
      throw new BadRequestException(
        'GSTR3B_VS_GSTR2B requires gstr3b_itc and gstr2b_itc components.',
      );
    }
    const toolResult = gstr3bVs2b({
      gstr3bItc: request.gstr3b_itc,
      gstr2bItc: request.gstr2b_itc,
    });
    const sheetName =
      request.output_sheet_name ??
      `Recon 3B-vs-2B ${request.period ?? ''}`.trim().slice(0, 31);

    const actions = map3bVs2bToActions({
      sheetName,
      period: request.period,
      gstin: request.gstin,
      result: toolResult.data,
      runAt,
      relativeTo: request.portal_file?.sheet_name ?? request.purchase_register?.sheet_name,
    });

    return {
      job_id: jobId,
      status: 'complete' as const,
      reconciliation_type: request.reconciliation_type,
      summary: {
        total_pr_rows: 0,
        total_portal_rows: 0,
        total_ims_rows: 0,
        exact_matched: 0,
        partial_matched: 0,
        credit_notes: 0,
        pr_only: 0,
        portal_only: 0,
        ims_rejected: 0,
        ims_pending: 0,
        ims_auto_accept: 0,
        rcm_flagged: 0,
        itc_matched: 0,
        itc_at_risk: toolResult.data.excessClaim,
        rcm_payable: 0,
        itc_ims_rejected: 0,
        itc_ims_pending: 0,
        gstr3b_vs_2b: toolResult.data,
      },
      rows: [],
      actions,
      confidence: toolResult.confidence,
      exceptions: toolResult.exceptions,
      source_refs: toolResult.sourceRefs,
      output_sheet_name: sheetName,
      audit_log_id: null,
    };
  }

  private reconcileInvoiceLevel(
    request: GstReconcileRequestDto,
    jobId: string,
    runAt: string,
  ) {
    if (!request.purchase_register?.data?.length) {
      throw new BadRequestException('purchase_register.data is required.');
    }

    const isImsOnly = request.reconciliation_type === 'IMS_VS_PR';
    if (!isImsOnly && !request.portal_file?.data?.length) {
      throw new BadRequestException('portal_file.data is required for this reconciliation type.');
    }

    const prSheet = request.purchase_register;
    const pr = parsePurchaseRegister(prSheet.data, {
      documentId: prSheet.sheet_name,
      headersRow: prSheet.headers_row,
      columnMapping: toColumnMapping(prSheet.column_mapping),
    });

    let portal: ReturnType<typeof parseGstr2b> = [];
    if (request.portal_file?.data?.length) {
      const pf = request.portal_file;
      const opts = {
        documentId: pf.sheet_name,
        columnMapping: toColumnMapping(pf.column_mapping),
      };
      if (
        request.reconciliation_type === 'PR_VS_GSTR2A' ||
        String(pf.file_type ?? '').toUpperCase() === 'GSTR2A'
      ) {
        portal = parseGstr2a(pf.data, opts);
      } else {
        portal = parseGstr2b(pf.data, opts);
      }
    }

    let ims = [] as ReturnType<typeof parseImsExport>;
    const useIms =
      Boolean(request.settings?.use_ims_data) ||
      request.reconciliation_type === 'IMS_VS_PR' ||
      Boolean(request.ims_data?.data?.length);

    if (request.ims_data?.data?.length) {
      ims = parseImsExport(request.ims_data.data, {
        documentId: request.ims_data.sheet_name,
        headersRow: request.ims_data.headers_row,
        columnMapping: toColumnMapping(request.ims_data.column_mapping),
      });
    }

    // IMS_VS_PR: use IMS as the "portal" side when no GSTR file
    if (isImsOnly && ims.length && !portal.length) {
      portal = ims.map((r) => ({ ...r }));
    }

    const settings: Partial<GstMatchSettings> = {
      ...DEFAULT_GST_MATCH_SETTINGS,
      ...settingsFromDto(request.settings),
      useImsData: useIms,
    };

    const matchResult = gstMatch({
      purchaseRegister: pr,
      gstr2b: portal,
      imsRows: useIms ? ims : undefined,
      settings,
    });

    const itcResult = itcCompute({
      resultRows: matchResult.data.resultRows,
    });

    const summary: ReconSummaryBlock = buildSummary(
      matchResult.data.resultRows,
      itcResult.data,
      pr.length,
      portal.length,
      ims.length,
    );

    const rows = mapResultRows(matchResult.data.resultRows);
    const sheetName =
      request.output_sheet_name ??
      `Recon ${request.reconciliation_type.replace(/_/g, ' ')} ${request.period ?? ''}`
        .trim()
        .slice(0, 31);

    const actions = mapToSheetActions({
      sheetName,
      period: request.period,
      gstin: request.gstin,
      reconType: request.reconciliation_type,
      summary,
      rows,
      runAt,
      relativeTo: prSheet.sheet_name,
    });

    return {
      job_id: jobId,
      status: 'complete' as const,
      reconciliation_type: request.reconciliation_type,
      summary,
      rows,
      actions,
      confidence: matchResult.confidence,
      exceptions: [...matchResult.exceptions, ...itcResult.exceptions],
      source_refs: matchResult.sourceRefs,
      output_sheet_name: sheetName,
      settings: matchResult.data.settings,
      audit_log_id: null,
    };
  }
}
