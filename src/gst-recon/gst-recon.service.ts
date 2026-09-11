import { BadRequestException, Injectable, Logger, Optional } from '@nestjs/common';
import { randomUUID } from 'crypto';
import {
  ColumnMapping,
  DEFAULT_GST_MATCH_SETTINGS,
  GstMatchMode,
  GstMatchSettings,
  NormalizedInvoiceRow,
} from '../domain-tools/types/domain-tool.types';
import { parsePurchaseRegister } from '../domain-tools/ingestion/purchase-register-parser';
import { parseSalesRegister } from '../domain-tools/ingestion/sales-register-parser';
import { hasInvoiceNumberColumn } from '../domain-tools/ingestion/portal-file-detector';
import { parseGstr2a, parseGstr2b } from '../domain-tools/ingestion/gstr2b-parser';
import { parseGstr1 } from '../domain-tools/ingestion/gstr1-parser';
import { parseImsExport } from '../domain-tools/ingestion/ims-parser';
import { gstMatch } from '../domain-tools/gst/gst-match.tool';
import { itcCompute } from '../domain-tools/gst/itc-compute.tool';
import { gstr3bVs2b } from '../domain-tools/gst/gstr3b-vs-2b.tool';
import { AuditService } from '../audit/audit.service';
import { GstReconcileRequestDto, SheetPayloadDto } from './gst-recon.dto';
import {
  buildSummary,
  map3bVs2bToActions,
  mapMissedBooksFlatToSheetActions,
  mapMissedBooksToSheetActions,
  mapPortalOnlyFlatToSheetActions,
  mapResultRows,
  mapToSheetActions,
  ReconSummaryBlock,
} from './gst-recon-to-actions.mapper';
import {
  extractPortalStatementGstin,
  flagCrossGstinRows,
  normalizeGstin,
  validateGstinFormat,
  validatePortalGstinMatchesClient,
} from './gstin-validator';
import { NormalizedRowWithClientGstinColumn } from './types';

const logger = new Logger('GstReconService');

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
  assign('clientGstin', (raw as { clientGstin?: string | number }).clientGstin);
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
  assign(
    'supplyCategory',
    (raw as { supplyCategory?: string | number }).supplyCategory,
  );
  assign(
    'placeOfSupply',
    (raw as { placeOfSupply?: string | number }).placeOfSupply,
  );
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

function resolveClientGstin(request: GstReconcileRequestDto): string {
  return normalizeGstin(request.client_gstin || request.gstin || '');
}

function invoiceToNormalizedRow(
  row: NormalizedInvoiceRow,
): NormalizedRowWithClientGstinColumn {
  if (row.taxableValue == null) {
    // Should only fire for a genuine data gap (e.g. no rate-slab column populated for
    // this row) — never for a column-detection failure. If this fires unexpectedly for
    // rows that clearly have a value in the source sheet, isRateSlabLayout /
    // deriveTaxableValueFromSlabRow are failing to match the sheet's headers.
    logger.warn(
      `Taxable value missing for books row ${row.sourceRowRef.rowOrLine} ` +
        `(invoice "${row.invoiceNumber || '(blank)'}", GSTIN "${row.gstin || '(blank)'}") — defaulting to 0.`,
    );
  }
  return {
    rowIndex:
      typeof row.sourceRowRef.rowOrLine === 'number'
        ? row.sourceRowRef.rowOrLine
        : Number(row.sourceRowRef.rowOrLine) || 0,
    counterpartyGstin: row.gstin || null,
    clientSideGstin: row.clientSideGstin ?? null,
    documentType:
      row.documentType === 'credit_note'
        ? 'credit_note'
        : row.documentType === 'debit_note'
          ? 'debit_note'
          : row.documentType === 'amended'
            ? 'amended_invoice'
            : 'invoice',
    invoiceNumber: row.normalizedInvoiceNumber,
    invoiceNumberRaw: row.invoiceNumber,
    invoiceDate: row.invoiceDate,
    taxableValue: row.taxableValue ?? 0,
    igst: row.igst,
    cgst: row.cgst,
    sgst: row.sgst,
    cess: 0,
    supplyCategory: row.supplyCategory,
    irn: row.irn,
  };
}

@Injectable()
export class GstReconService {
  constructor(@Optional() private readonly auditService?: AuditService) {}

  async reconcile(request: GstReconcileRequestDto) {
    const jobId = `recon_${randomUUID().replace(/-/g, '').slice(0, 12)}`;
    const runAt = new Date().toISOString();

    if (request.reconciliation_type === 'GSTR3B_VS_GSTR2B') {
      return this.reconcile3bVs2b(request, jobId, runAt);
    }

    return this.reconcileInvoiceLevel(request, jobId, runAt);
  }

  async updateAuditOutcome(auditLogId: string, outcome: 'applied' | 'rejected') {
    if (!this.auditService) {
      throw new BadRequestException('Audit service unavailable.');
    }
    const ok = await this.auditService.updateGstReconOutcome(auditLogId, outcome);
    if (!ok) {
      throw new BadRequestException('GST recon audit entry not found.');
    }
    return { audit_log_id: auditLogId, outcome };
  }

  private async reconcile3bVs2b(request: GstReconcileRequestDto, jobId: string, runAt: string) {
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

    const clientGstin = resolveClientGstin(request);
    const actions = map3bVs2bToActions({
      sheetName,
      period: request.period,
      gstin: clientGstin || request.gstin,
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
        cross_gstin_exception_count: 0,
        matched_exact: 0,
        matched_fallback: 0,
        mismatch_blank_gstin: 0,
        mismatch_blank_gstin_likely_matched: 0,
        mismatch_blank_taxable_value: 0,
        mismatch_ambiguous_rate_slab: 0,
        mismatch_amount: 0,
        mismatch_date: 0,
        mismatch_genuinely_missing: 0,
        gstin_mismatch_count: 0,
      },
      rows: [],
      actions,
      confidence: toolResult.confidence,
      exceptions: toolResult.exceptions,
      source_refs: toolResult.sourceRefs,
      output_sheet_name: sheetName,
      audit_log_id: null as string | null,
      client_gstin: clientGstin || null,
      cross_gstin_exception_count: 0,
    };
  }

  private async reconcileInvoiceLevel(
    request: GstReconcileRequestDto,
    jobId: string,
    runAt: string,
  ) {
    const isSales = request.reconciliation_type === 'SALES_VS_GSTR1';
    const booksSheet = request.books_register ?? request.purchase_register;
    if (!booksSheet?.data?.length) {
      throw new BadRequestException(
        isSales
          ? 'books_register.data (sales register) is required.'
          : 'purchase_register.data is required.',
      );
    }

    const isImsOnly = request.reconciliation_type === 'IMS_VS_PR';
    const hasPrimaryPortal = Boolean(request.portal_file?.data?.length);
    const hasPortal2a = Boolean(request.portal_file_2a?.data?.length);
    if (!isImsOnly && !hasPrimaryPortal && !hasPortal2a) {
      throw new BadRequestException('portal_file.data is required for this reconciliation type.');
    }

    const missedBooksOnly = Boolean(request.missed_books_only);
    const clientGstinRaw = resolveClientGstin(request);
    let clientGstin = '';
    if (clientGstinRaw) {
      const gstinCheck = validateGstinFormat(clientGstinRaw);
      if (!gstinCheck.ok) {
        throw new BadRequestException(
          gstinCheck.blockingError ??
            'client_gstin is required and must be a valid GSTIN for reconciliation.',
        );
      }
      clientGstin = gstinCheck.normalizedGstin!;
    } else if (!missedBooksOnly) {
      throw new BadRequestException(
        'client_gstin is required and must be a valid GSTIN for reconciliation.',
      );
    }

    if (clientGstin && request.portal_file?.data?.length) {
      const portalGstin = extractPortalStatementGstin(
        request.portal_file.data,
        isSales ? 'sales_vs_gstr1' : 'purchase_vs_2b',
      );
      if (portalGstin) {
        const portalCheck = validatePortalGstinMatchesClient(clientGstin, portalGstin);
        if (!portalCheck.ok) {
          throw new BadRequestException(portalCheck.blockingError);
        }
      }
    }
    if (clientGstin && request.portal_file_2a?.data?.length) {
      const portalGstin = extractPortalStatementGstin(
        request.portal_file_2a.data,
        'purchase_vs_2a',
      );
      if (portalGstin) {
        const portalCheck = validatePortalGstinMatchesClient(clientGstin, portalGstin);
        if (!portalCheck.ok) {
          throw new BadRequestException(portalCheck.blockingError);
        }
      }
    }

    const booksColumnMapping = toColumnMapping(booksSheet.column_mapping);
    let books = isSales
      ? parseSalesRegister(booksSheet.data, {
          documentId: booksSheet.sheet_name,
          headersRow: booksSheet.headers_row,
          columnMapping: booksColumnMapping,
        })
      : parsePurchaseRegister(booksSheet.data, {
          documentId: booksSheet.sheet_name,
          headersRow: booksSheet.headers_row,
          columnMapping: booksColumnMapping,
        });

    // Checked once per sheet (not per row) — enables the Pass 1b GSTIN+date+amount
    // fallback match when the books export has no Invoice Number column at all.
    const booksHeaderRowIndex = Math.max(0, (booksSheet.headers_row ?? 1) - 1);
    const booksHeaderRow = (booksSheet.data?.[booksHeaderRowIndex] ?? []) as unknown[];
    const booksHasInvoiceNumberColumn =
      Boolean(booksColumnMapping?.invoiceNo) || hasInvoiceNumberColumn(booksHeaderRow);

    // Period scoping: when the prompt yielded a confident period, filter books rows to it
    // BEFORE anything downstream (cross-GSTIN flagging, matching) ever sees the out-of-period
    // rows. Portal rows are filtered the same way once the portal array is fully assembled below.
    const periodStart = request.period_start;
    const periodEnd = request.period_end;
    const hasPeriodFilter = Boolean(periodStart && periodEnd);
    const inPeriod = (d: string) => Boolean(d) && d >= periodStart! && d <= periodEnd!;
    const booksCountBeforePeriodFilter = books.length;
    if (hasPeriodFilter) {
      books = books.filter((r) => inPeriod(r.invoiceDate));
    }

    // Cross-GSTIN flagging on books rows
    const asNorm = books.map(invoiceToNormalizedRow);
    const { clean, crossGstin } = clientGstin
      ? flagCrossGstinRows(asNorm, clientGstin)
      : { clean: asNorm, crossGstin: [] as typeof asNorm };
    const cleanRowIndexes = new Set(clean.map((r) => r.rowIndex));
    const cleanBooks = books.filter((b) => {
      const idx =
        typeof b.sourceRowRef.rowOrLine === 'number'
          ? b.sourceRowRef.rowOrLine
          : Number(b.sourceRowRef.rowOrLine) || 0;
      return cleanRowIndexes.has(idx);
    });

    let portal: NormalizedInvoiceRow[] = [];
    const portalLabels: string[] = [];
    if (request.portal_file?.data?.length) {
      const pf = request.portal_file;
      const opts = {
        documentId: pf.sheet_name,
        columnMapping: toColumnMapping(pf.column_mapping),
        headersRow: pf.headers_row,
      };
      const fileType = String(pf.file_type ?? '').toUpperCase();
      if (isSales || fileType === 'GSTR1') {
        portal = parseGstr1(pf.data, opts);
        portalLabels.push('GSTR-1');
      } else if (
        request.reconciliation_type === 'PR_VS_GSTR2A' ||
        fileType === 'GSTR2A'
      ) {
        portal = parseGstr2a(pf.data, opts);
        portalLabels.push('GSTR-2A');
      } else {
        portal = parseGstr2b(pf.data, opts);
        portalLabels.push('GSTR-2B');
      }
    }
    if (!isSales && request.portal_file_2a?.data?.length) {
      const pf2a = request.portal_file_2a;
      const extra = parseGstr2a(pf2a.data, {
        documentId: pf2a.sheet_name,
        columnMapping: toColumnMapping(pf2a.column_mapping),
        headersRow: pf2a.headers_row,
      });
      portal = [...portal, ...extra];
      if (!portalLabels.includes('GSTR-2A')) portalLabels.push('GSTR-2A');
    }

    const portalCountBeforePeriodFilter = portal.length;
    if (hasPeriodFilter) {
      portal = portal.filter((r) => inPeriod(r.invoiceDate));
    }

    const periodLabel = request.period_label || `${periodStart} to ${periodEnd}`;
    const periodApplied = hasPeriodFilter ? { label: periodLabel, start: periodStart!, end: periodEnd! } : undefined;
    const booksLabelForPeriodMsg = isSales ? 'Sales Register' : 'Purchase Register';
    const portalLabelForPeriodMsg = portalLabels.join(' / ') || (isSales ? 'GSTR-1' : 'GSTR-2B');
    let periodZeroMessage: string | undefined;
    if (hasPeriodFilter && booksCountBeforePeriodFilter > 0 && books.length === 0) {
      periodZeroMessage = `No ${booksLabelForPeriodMsg} rows found for ${periodLabel} — check the date column or try a different period.`;
    } else if (hasPeriodFilter && portalCountBeforePeriodFilter > 0 && portal.length === 0) {
      periodZeroMessage = `No ${portalLabelForPeriodMsg} rows found for ${periodLabel} — check the date column or try a different period.`;
    }

    if (periodZeroMessage) {
      const zeroSummary: ReconSummaryBlock = {
        total_pr_rows: books.length,
        total_portal_rows: portal.length,
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
        itc_at_risk: 0,
        rcm_payable: 0,
        itc_ims_rejected: 0,
        itc_ims_pending: 0,
        cross_gstin_exception_count: 0,
        matched_exact: 0,
        matched_fallback: 0,
        mismatch_blank_gstin: 0,
        mismatch_blank_gstin_likely_matched: 0,
        mismatch_blank_taxable_value: 0,
        mismatch_ambiguous_rate_slab: 0,
        mismatch_amount: 0,
        mismatch_date: 0,
        mismatch_genuinely_missing: 0,
        gstin_mismatch_count: 0,
      };
      return {
        job_id: jobId,
        status: 'complete' as const,
        reconciliation_type: request.reconciliation_type,
        actionType: 'GST_RECON_RESULT' as const,
        client_gstin: clientGstin || null,
        client_name: request.client_name ?? null,
        summary: zeroSummary,
        rows: [],
        actions: [],
        confidence: 1,
        exceptions: [],
        source_refs: [],
        output_sheet_name: request.output_sheet_name ?? '',
        audit_log_id: null as string | null,
        cross_gstin_exception_count: 0,
        missed_books_only: missedBooksOnly,
        portal_label: portalLabelForPeriodMsg,
        books_label: booksLabelForPeriodMsg,
        period_applied: periodApplied,
        period_zero_message: periodZeroMessage,
      };
    }

    let ims = [] as ReturnType<typeof parseImsExport>;
    const useIms =
      !isSales &&
      (Boolean(request.settings?.use_ims_data) ||
        request.reconciliation_type === 'IMS_VS_PR' ||
        Boolean(request.ims_data?.data?.length));

    if (request.ims_data?.data?.length) {
      ims = parseImsExport(request.ims_data.data, {
        documentId: request.ims_data.sheet_name,
        headersRow: request.ims_data.headers_row,
        columnMapping: toColumnMapping(request.ims_data.column_mapping),
      });
    }

    if (isImsOnly && ims.length && !portal.length) {
      portal = ims.map((r) => ({ ...r }));
    }

    const settings: Partial<GstMatchSettings> = {
      ...DEFAULT_GST_MATCH_SETTINGS,
      ...settingsFromDto(request.settings),
      useImsData: useIms,
      detectRcm: missedBooksOnly
        ? false
        : isSales
          ? false
          : (request.settings?.detect_rcm ?? true),
    };

    const mode: GstMatchMode = isSales ? 'sales' : 'purchase';
    const matchResult = gstMatch({
      purchaseRegister: cleanBooks,
      gstr2b: portal,
      imsRows: useIms ? ims : undefined,
      settings,
      mode,
      booksHasInvoiceNumberColumn,
    });

    const itcResult = itcCompute({
      resultRows: matchResult.data.resultRows,
    });

    const summary: ReconSummaryBlock = {
      ...buildSummary(
        matchResult.data.resultRows,
        itcResult.data,
        books.length,
        portal.length,
        ims.length,
      ),
      cross_gstin_exception_count: crossGstin.length,
    };

    const rows = mapResultRows(matchResult.data.resultRows);
    // Append cross-GSTIN exceptions as flagged rows for the report
    for (const x of crossGstin) {
      rows.push({
        pr_ref: `${booksSheet.sheet_name}:row_${x.rowIndex}`,
        portal_ref: null,
        status: 'CROSS_GSTIN',
        pass: null,
        confidence: 1,
        difference: `Books GSTIN ${x.clientSideGstin} ≠ client ${clientGstin}`,
        diff_type: 'CROSS_GSTIN',
        itc_amount: 0,
        ims_status: null,
        rcm_flag: false,
        invoice_number: x.invoiceNumberRaw,
        gstin: x.counterpartyGstin,
        vendor_name: null,
        mismatch_reason: null,
        explanation: null,
      });
    }

    const portalLabel = portalLabels.join(' / ') || (isSales ? 'GSTR-1' : 'GSTR-2B');
    const booksLabel = isSales ? 'Sales Register' : 'Purchase Register';
    const sheetName =
      request.output_sheet_name ??
      (missedBooksOnly
        ? `Missed vs ${portalLabel}`.slice(0, 31)
        : `Recon ${request.reconciliation_type.replace(/_/g, ' ')} ${request.period ?? ''}`
            .trim()
            .slice(0, 31));

    const layout = request.layout ?? 'categorized';
    const missedBooksParams = {
      sheetName,
      portalLabel,
      booksLabel,
      runAt,
      relativeTo: booksSheet.sheet_name,
      resultRows: matchResult.data.resultRows,
    };
    const actions = !missedBooksOnly
      ? mapToSheetActions({
          sheetName,
          period: request.period,
          gstin: clientGstin,
          clientName: request.client_name,
          financialYear: request.financial_year,
          operatorName: request.operator_name,
          firmName: request.firm_name,
          reconType: request.reconciliation_type,
          summary,
          rows,
          runAt,
          relativeTo: booksSheet.sheet_name,
          isSales,
        })
      : layout === 'books_flat'
        ? mapMissedBooksFlatToSheetActions(missedBooksParams)
        : layout === 'portal_flat'
          ? mapPortalOnlyFlatToSheetActions(missedBooksParams)
          : mapMissedBooksToSheetActions(missedBooksParams);

    let auditLogId: string | null = null;
    if (this.auditService) {
      try {
        auditLogId = await this.auditService.createGstReconAudit({
          jobId,
          clientGstin,
          reconType: request.reconciliation_type,
          period: request.period,
          clientName: request.client_name,
          sourceSheets: {
            books: booksSheet.sheet_name,
            portal: [request.portal_file?.sheet_name, request.portal_file_2a?.sheet_name]
              .filter(Boolean)
              .join(', '),
          },
          crossGstinExceptionCount: crossGstin.length,
          matchingSettings: settings as Record<string, unknown>,
          operatorName: request.operator_name,
          firmName: request.firm_name,
        });
      } catch {
        auditLogId = null;
      }
    }

    return {
      job_id: jobId,
      status: 'complete' as const,
      reconciliation_type: request.reconciliation_type,
      actionType: 'GST_RECON_RESULT' as const,
      client_gstin: clientGstin,
      client_name: request.client_name ?? null,
      summary,
      rows,
      actions,
      confidence: matchResult.confidence,
      exceptions: [
        ...matchResult.exceptions,
        ...itcResult.exceptions,
        ...(crossGstin.length
          ? [
              {
                code: 'GST_CROSS_GSTIN_ROWS',
                severity: 'flag' as const,
                message: `${crossGstin.length} books row(s) have a different client-side GSTIN than ${clientGstin}`,
                affectedRows: crossGstin.map((r) => r.rowIndex),
              },
            ]
          : []),
      ],
      source_refs: matchResult.sourceRefs,
      output_sheet_name: sheetName,
      settings: matchResult.data.settings,
      audit_log_id: auditLogId,
      cross_gstin_exception_count: crossGstin.length,
      missed_books_only: missedBooksOnly,
      portal_label: portalLabel,
      books_label: booksLabel,
      period_applied: periodApplied,
    };
  }
}
