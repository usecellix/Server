import { Injectable } from '@nestjs/common';
import {
  GstReconIntent,
  GstReconIntentPayload,
  ReconOrchestratorResult,
} from './orchestrator.types';
import { resolveMissingContext } from './context-resolver';
import {
  ambiguousSheetMessage,
  resolveSheetsForRecon,
  SHEET_MESSAGES,
} from './sheet-detector';
import { RegisterSignature, ReconType } from './types';
import {
  proposeClientGstinFromBooks,
  validateGstinFormat,
} from './gstin-validator';
import { GstReconService } from './gst-recon.service';
import { GstReconcileRequestDto } from './gst-recon.dto';

function inferPurchaseSubtype(payload: GstReconIntentPayload): ReconType {
  const raw = `${payload.extractedPeriod ?? ''} ${JSON.stringify(payload)}`.toLowerCase();
  if (/\b2a\b|gstr[\s\-]?2a/.test(raw)) return 'purchase_vs_2a';
  return 'purchase_vs_2b';
}

function buildPartialContext(
  payload: GstReconIntentPayload,
  reconType: ReconType,
) {
  return {
    clientGstin: payload.extractedGstin,
    taxPeriod: payload.extractedPeriod,
    financialYear: payload.extractedFinancialYear,
    client: payload.extractedClientName
      ? { id: '', name: payload.extractedClientName }
      : undefined,
    reconType,
  };
}

function buildReconSheetName(
  clientName: string | undefined,
  period: string,
  reconType: ReconType,
): string {
  const typeLabel =
    reconType === 'sales_vs_gstr1'
      ? 'Sales'
      : reconType === 'purchase_vs_2a'
        ? 'Purchase-2A'
        : 'Purchase';
  const base = `GST Recon — ${clientName ?? 'Client'} — ${period} — ${typeLabel}`;
  return base.slice(0, 31);
}

/** If row 1 is a GSTIN metadata banner, headers are on row 2. */
function inferHeadersRow(grid: unknown[][]): number {
  if (!grid?.length) return 1;
  const joined = (grid[0] ?? [])
    .map((c) =>
      String(c ?? '')
        .trim()
        .toLowerCase(),
    )
    .join(' ');
  if (
    joined.includes('gstin of the taxpayer') ||
    joined.includes('gstin of registered') ||
    joined.includes('taxpayer gstin')
  ) {
    return 2;
  }
  return 1;
}

@Injectable()
export class GstReconOrchestrator {
  constructor(private readonly gstReconService: GstReconService) {}

  /**
   * Conversational orchestration (spec §6). Matching logic stays in GstReconService.
   */
  async runConversationalRecon(
    intent: GstReconIntent,
    payload: GstReconIntentPayload,
    openWorkbookSheets: { sheetName: string; headers: string[] }[],
    sheetData?: Record<string, unknown[][]>,
  ): Promise<ReconOrchestratorResult> {
    const reconType: ReconType =
      intent === 'GST_RECON_PURCHASE'
        ? inferPurchaseSubtype(payload)
        : 'sales_vs_gstr1';

    const partialContext = buildPartialContext(payload, reconType);
    const missing = resolveMissingContext(partialContext);
    if (missing.length > 0) {
      return {
        kind: 'needs_input',
        prompts: missing.map((m) => m.chatPrompt),
      };
    }

    const gstinCheck = validateGstinFormat(partialContext.clientGstin!);
    if (!gstinCheck.ok) {
      return { kind: 'error', message: gstinCheck.blockingError! };
    }

    const needed: RegisterSignature[] =
      reconType === 'sales_vs_gstr1'
        ? ['sales_register', 'gstr_1']
        : reconType === 'purchase_vs_2a'
          ? ['purchase_register', 'gstr_2a']
          : ['purchase_register', 'gstr_2b'];

    const resolved = resolveSheetsForRecon(openWorkbookSheets, needed);
    for (const sig of needed) {
      if (resolved[sig].status === 'not_found') {
        return { kind: 'chat_reply', message: SHEET_MESSAGES[sig].notFound };
      }
      if (resolved[sig].status === 'ambiguous') {
        return {
          kind: 'chat_reply',
          message: ambiguousSheetMessage(sig, resolved[sig].found),
        };
      }
    }

    if (!sheetData) {
      return {
        kind: 'needs_sheet_data',
        booksSheet: resolved[needed[0]].found[0].sheetName,
        portalSheet: resolved[needed[1]].found[0].sheetName,
        reconType,
        clientGstin: gstinCheck.normalizedGstin!,
        period: partialContext.taxPeriod!,
        clientName: partialContext.client?.name,
        financialYear: partialContext.financialYear,
      };
    }

    const booksSheetName = resolved[needed[0]].found[0].sheetName;
    const portalSheetName = resolved[needed[1]].found[0].sheetName;
    const booksRows = sheetData[booksSheetName];
    const portalRows = sheetData[portalSheetName];
    if (!booksRows?.length || !portalRows?.length) {
      return {
        kind: 'error',
        message: 'Could not read row data for the resolved sheets.',
      };
    }

    const dtoType =
      reconType === 'sales_vs_gstr1'
        ? 'SALES_VS_GSTR1'
        : reconType === 'purchase_vs_2a'
          ? 'PR_VS_GSTR2A'
          : 'PR_VS_GSTR2B';

    const booksHeaderRow = inferHeadersRow(booksRows);
    const portalHeaderRow = inferHeadersRow(portalRows);

    const request: GstReconcileRequestDto = {
      reconciliation_type: dtoType,
      client_gstin: gstinCheck.normalizedGstin!,
      period: partialContext.taxPeriod,
      financial_year: partialContext.financialYear,
      client_name: partialContext.client?.name,
      purchase_register: {
        sheet_name: booksSheetName,
        data: booksRows,
        headers_row: booksHeaderRow,
      },
      books_register: {
        sheet_name: booksSheetName,
        data: booksRows,
        headers_row: booksHeaderRow,
      },
      portal_file: {
        sheet_name: portalSheetName,
        data: portalRows,
        headers_row: portalHeaderRow,
        file_type:
          reconType === 'sales_vs_gstr1'
            ? 'GSTR1'
            : reconType === 'purchase_vs_2a'
              ? 'GSTR2A'
              : 'GSTR2B',
      },
      output_sheet_name: buildReconSheetName(
        partialContext.client?.name,
        partialContext.taxPeriod!,
        reconType,
      ),
    };

    try {
      const result = await this.gstReconService.reconcile(request);
      return {
        kind: 'action_payload',
        summary: result.summary,
        matchResults: result.rows,
        actions: result.actions,
        suggestedSheetName: result.output_sheet_name,
        jobId: result.job_id,
        clientGstin: gstinCheck.normalizedGstin!,
        reconType,
        crossGstinExceptionCount: result.cross_gstin_exception_count ?? 0,
        auditLogId: result.audit_log_id,
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return { kind: 'error', message };
    }
  }

  /** Helper for proposing GSTIN from books when sheets already read. */
  proposeGstinFromGrid(grid: unknown[][]): string | undefined {
    // Lightweight scan of first data column that looks like client GSTIN — optional
    const rows = grid.slice(0, 30).map((r) => ({
      clientSideGstin: Array.isArray(r) ? String(r[0] ?? '') : '',
    }));
    return proposeClientGstinFromBooks(rows);
  }
}
