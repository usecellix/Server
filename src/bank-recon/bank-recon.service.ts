import { BadRequestException, Injectable } from '@nestjs/common';
import { randomUUID } from 'crypto';
import { sanitizeExcelSheetName } from '../excel-ai/utils/sheet-name.util';
import { DEFAULT_BANK_MATCH_SETTINGS, BankMatchSettings } from './types';
import { BankReconRelayoutRequestDto, BankReconcileRequestDto } from './bank-recon.dto';
import { runBankReconEngine } from './bank-recon.engine';
import { getBankLayoutCache, storeBankLayoutCache } from './layout-cache';
import {
  BANK_REASON_LABEL,
  BOOKS_REASON_LABEL,
  BankSheetModel,
  bankSheetName,
  NOT_CLEARED_SHEET,
  TIMING_REASONS,
  mapBankReconToActions,
  reportableBank,
} from './bank-recon-to-actions.mapper';

function settingsFromDto(dto: BankReconcileRequestDto['settings']): Partial<BankMatchSettings> {
  if (!dto) return {};
  const out: Partial<BankMatchSettings> = {};
  if (dto.clearing_window_days !== undefined) out.clearingWindowDays = dto.clearing_window_days;
  if (dto.tight_window_days !== undefined) out.tightWindowDays = dto.tight_window_days;
  if (dto.amount_tolerance !== undefined) out.amountTolerance = dto.amount_tolerance;
  if (dto.narration_threshold !== undefined) out.narrationThreshold = dto.narration_threshold;
  if (dto.max_group_size !== undefined) out.maxGroupSize = dto.max_group_size;
  if (dto.enable_group_match !== undefined) out.enableGroupMatch = dto.enable_group_match;
  return out;
}

/** A reconciling item the client renders as a clickable "jump to row" list. */
export interface BankReconRowDto {
  side: 'BOOKS' | 'BANK';
  sheet_name: string;
  row: number;
  date: string;
  direction: 'IN' | 'OUT';
  amount: number;
  narration: string;
  reference: string;
  reason: string;
  reason_label: string;
  explanation: string;
}

@Injectable()
export class BankReconService {
  reconcile(request: BankReconcileRequestDto) {
    const jobId = `bankrecon_${randomUUID().replace(/-/g, '').slice(0, 12)}`;
    const runAt = new Date().toISOString();
    const casual = request.casual ?? true;
    const layout = request.layout ?? 'categorized';

    if (!request.books_ledger?.data?.length) {
      throw new BadRequestException('books_ledger.data (the bank ledger from your books) is required.');
    }
    if (!request.bank_statement?.data?.length) {
      throw new BadRequestException('bank_statement.data is required.');
    }

    const out = runBankReconEngine({
      books: {
        sheetName: request.books_ledger.sheet_name,
        data: request.books_ledger.data,
        headersRow: request.books_ledger.headers_row,
        fileType: request.books_ledger.file_type,
      },
      bank: {
        sheetName: request.bank_statement.sheet_name,
        data: request.bank_statement.data,
        headersRow: request.bank_statement.headers_row,
        fileType: request.bank_statement.file_type,
      },
      priorBrs: request.prior_brs?.data?.length
        ? { sheetName: request.prior_brs.sheet_name, data: request.prior_brs.data }
        : undefined,
      periodStart: request.period_start,
      periodEnd: request.period_end,
      booksOpeningBalance: request.books_opening_balance,
      bankOpeningBalance: request.bank_opening_balance,
      accountNumber: request.casual === false ? request.account_number : undefined,
      settings: { ...DEFAULT_BANK_MATCH_SETTINGS, ...settingsFromDto(request.settings) },
    });

    const validationErrors = [...out.booksValidation.errors, ...out.bankValidation.errors];
    const warnings = [...out.booksValidation.warnings, ...out.bankValidation.warnings].map((w) => w.message);

    const base = {
      job_id: jobId,
      status: 'complete' as const,
      actionType: 'BANK_RECON_RESULT' as const,
      client_name: request.client_name ?? null,
      layout,
      period_applied: {
        label: request.period_label || `${out.periodStart} to ${out.periodEnd}`,
        start: out.periodStart,
        end: out.periodEnd,
      },
      carry_forward_items: out.carryForwardCount,
      validation: { errors: validationErrors, warnings, notes: out.notes },
    };

    if (out.periodZeroMessage) {
      return {
        ...base,
        summary: emptySummary(),
        rows: [] as BankReconRowDto[],
        actions: [],
        brs: null,
        output_sheet_name: '',
        layout_cache_id: null as string | null,
        period_zero_message: out.periodZeroMessage,
      };
    }

    // Gated (full) runs must not produce a misleading reconciliation from data that fails
    // the balance tie-out; casual runs surface the same problems as warnings instead.
    if (!casual && validationErrors.length) {
      throw new BadRequestException(
        `I can't reconcile yet — the data failed validation:\n\n${validationErrors.map((e) => `• ${e.message}`).join('\n')}`,
      );
    }

    const timingOnly = layout === 'books_flat' && request.books_filter === 'timing';
    const sheetName = request.output_sheet_name
      ? sanitizeExcelSheetName(request.output_sheet_name)
      : timingOnly
        ? NOT_CLEARED_SHEET
        : bankSheetName(layout);

    const model: BankSheetModel = {
      result: out.result,
      brs: out.brs,
      periodLabel: base.period_applied.label,
      periodStart: out.periodStart,
      periodEnd: out.periodEnd,
      runAt,
      relativeTo: request.books_ledger.sheet_name,
      booksSheet: out.booksParsed.txns[0]?.sourceLocation.sheet ?? request.books_ledger.sheet_name,
      bankSheet: out.bankParsed.txns[0]?.sourceLocation.sheet ?? request.bank_statement.sheet_name,
    };

    // The "not cleared" view lists only timing items; the cached full result keeps everything for relayout.
    const viewModel: BankSheetModel = timingOnly
      ? {
          ...model,
          result: {
            ...model.result,
            unmatchedBooks: model.result.unmatchedBooks.filter((u) => TIMING_REASONS.includes(u.reason)),
          },
        }
      : model;
    const actions = mapBankReconToActions(layout, viewModel, sheetName);
    const layoutCacheId = storeBankLayoutCache({
      result: out.result,
      brs: out.brs,
      periodLabel: model.periodLabel,
      periodStart: model.periodStart,
      periodEnd: model.periodEnd,
      relativeTo: model.relativeTo,
      layout,
      sheetName,
      booksSheet: model.booksSheet,
      bankSheet: model.bankSheet,
    });

    const rows = toRows(viewModel);
    return {
      ...base,
      summary: buildSummary(out),
      rows,
      actions,
      brs: out.brs,
      output_sheet_name: sheetName,
      layout_cache_id: layoutCacheId,
    };
  }

  /** Re-map an already-completed run into a different layout — no parsing, no matching. */
  relayout(request: BankReconRelayoutRequestDto) {
    const cached = getBankLayoutCache(request.layout_cache_id);
    if (!cached) {
      throw new BadRequestException(
        'That reconciliation result is no longer available — please run the bank reconciliation again.',
      );
    }
    const sheetName = request.output_sheet_name
      ? sanitizeExcelSheetName(request.output_sheet_name)
      : bankSheetName(request.layout);
    const model: BankSheetModel = {
      result: cached.result,
      brs: cached.brs,
      periodLabel: cached.periodLabel,
      periodStart: cached.periodStart,
      periodEnd: cached.periodEnd,
      runAt: new Date().toISOString(),
      relativeTo: cached.relativeTo,
      booksSheet: cached.booksSheet,
      bankSheet: cached.bankSheet,
    };
    return {
      actions: mapBankReconToActions(request.layout, model, sheetName),
      output_sheet_name: sheetName,
      layout: request.layout,
    };
  }
}

function emptySummary() {
  return {
    books_rows: 0,
    bank_rows: 0,
    matched: 0,
    matched_group: 0,
    matched_books_rows: 0,
    reversal_pairs: 0,
    books_only: 0,
    bank_only: 0,
    low_confidence_rows: 0,
    books_reasons: {} as Record<string, number>,
    bank_reasons: {} as Record<string, number>,
  };
}

function buildSummary(out: ReturnType<typeof runBankReconEngine>) {
  const { result } = out;
  const count = <T extends string>(items: Array<{ reason: T }>) => {
    const m: Record<string, number> = {};
    for (const i of items) m[i.reason] = (m[i.reason] ?? 0) + 1;
    return m;
  };
  const reportBank = result.unmatchedBank.filter(
    (u) => u.txn.txnDate >= out.periodStart && u.txn.txnDate <= out.periodEnd,
  );
  const lowConf = [
    ...result.unmatchedBooks.map((u) => u.txn),
    ...reportBank.map((u) => u.txn),
    ...result.matches.flatMap((m) => [...m.books, ...m.bank]),
  ].filter((t) => t.parseConfidence < 0.9).length;
  return {
    books_rows: out.booksRowCount,
    bank_rows: out.bankRowCount,
    matched: result.matches.filter((m) => m.kind === 'ONE_TO_ONE').length,
    matched_group: result.matches.filter((m) => m.kind === 'GROUP').length,
    // Ledger entries (not matches) that were reconciled — carry-forward items are not part of this period's ledger.
    matched_books_rows: result.matches
      .filter((m) => m.kind !== 'REVERSAL')
      .reduce((n, m) => n + m.books.filter((b) => b.voucherType !== 'Carry-forward').length, 0),
    reversal_pairs: result.matches.filter((m) => m.kind === 'REVERSAL').length,
    books_only: result.unmatchedBooks.length,
    bank_only: reportBank.length,
    low_confidence_rows: lowConf,
    books_reasons: count(result.unmatchedBooks),
    bank_reasons: count(reportBank),
  };
}

function toRows(model: BankSheetModel): BankReconRowDto[] {
  const mk = (
    side: 'BOOKS' | 'BANK',
    t: BankSheetModel['result']['unmatchedBooks'][number]['txn'],
    reason: string,
    label: string,
    explanation: string,
  ): BankReconRowDto => ({
    side,
    sheet_name: t.sourceLocation.sheet,
    row: t.sourceLocation.row,
    date: t.txnDate,
    direction: t.direction,
    amount: t.amount,
    narration: t.narration,
    reference: [t.instrumentNo, t.referenceNo].filter(Boolean).join(' / '),
    reason,
    reason_label: label,
    explanation,
  });
  return [
    ...model.result.unmatchedBooks.map((u) => mk('BOOKS', u.txn, u.reason, BOOKS_REASON_LABEL[u.reason], u.explanation)),
    ...reportableBank(model).map((u) => mk('BANK', u.txn, u.reason, BANK_REASON_LABEL[u.reason], u.explanation)),
  ];
}
