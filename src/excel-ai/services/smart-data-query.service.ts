import { Injectable, Logger } from '@nestjs/common';
import { WorkbookContext } from '../../types/cellix.types';
import {
  buildDataQuerySystemPrompt,
  buildDataQueryUserMessage,
} from '../prompts/data-query-system-prompt';
import { sliceRelevantColumns } from '../utils/column-slicer.util';
import { OpenRouterService } from './openrouter.service';

export type SmartDataQueryEmit = (event: string, data: Record<string, unknown>) => void;

/**
 * "How many sheets are in this workbook?" / "What sheets does this workbook
 * have?" / "List the sheets" — a WORKBOOK-STRUCTURE question, not a data
 * lookup. This route is deliberately scoped to the single active sheet's
 * data (see handleDataQueryRoute's own comment), so a workbook-wide sheet
 * question routed here always failed with "I could not find any sheet data"
 * whenever the active sheet happened to be empty — even though the answer
 * needs no sheet DATA at all, only the sheet list already sitting in
 * `workbookContext`. Live repro: right after creating a new blank sheet
 * (making it active), "How many sheets are in this workbook?" hard-failed
 * on a 4-sheet workbook. TASKS.md #256.
 */
const WORKBOOK_SHEET_COUNT_PATTERN =
  /\bhow\s+many\s+(?:sheets?|tabs?)\b|\bwhat\s+(?:sheets?|tabs?)\b|\blist\s+(?:the\s+|all\s+)?(?:sheets?|tabs?)\b|\bsheet\s+names?\b|\bnames?\s+of\s+(?:the\s+)?sheets?\b/i;

@Injectable()
export class SmartDataQueryService {
  private readonly logger = new Logger(SmartDataQueryService.name);

  constructor(private readonly openRouter: OpenRouterService) {}

  /**
   * Answer a read-only data query using column-sliced sheet data and the MEDIUM LLM tier.
   */
  async handleQuery(
    message: string,
    sheetData: unknown[][],
    workbookContext: WorkbookContext | undefined,
    activeSheetName: string | undefined,
    emit: SmartDataQueryEmit,
  ): Promise<string> {
    if (WORKBOOK_SHEET_COUNT_PATTERN.test(message)) {
      const answer = this.answerSheetCountQuestion(workbookContext);
      if (answer) return answer;
    }

    const sliceResult = sliceRelevantColumns(
      message,
      workbookContext,
      sheetData,
      activeSheetName,
    );

    if (!sliceResult.sheets.length || !sliceResult.sheets[0].rows.length) {
      this.logger.warn('SmartDataQuery: no sheet data available');
      return 'I could not find any sheet data to answer your question. Please make sure a sheet with data is active.';
    }

    const sheet = sliceResult.sheets[0];

    this.logger.log(
      `SmartDataQuery: sheet=${sheet.sheetName} cols=${sheet.headers.join(',')} rows=${sheet.totalRows}`,
    );

    const MAX_COLUMNS_IN_MESSAGE = 4;
    const columnSummary =
      sheet.headers.length > MAX_COLUMNS_IN_MESSAGE
        ? `${sheet.headers.slice(0, MAX_COLUMNS_IN_MESSAGE).join(', ')} + ${
            sheet.headers.length - MAX_COLUMNS_IN_MESSAGE
          } more columns`
        : sheet.headers.join(', ');

    emit('thinking', {
      message: `Reading ${sheet.sheetName} (${sheet.totalRows} rows) — analyzing ${columnSummary}`,
    });

    const systemPrompt = buildDataQuerySystemPrompt();
    const userMessage = buildDataQueryUserMessage(message, sheet, workbookContext);

    try {
      const answer = await this.openRouter.complete({
        systemPrompt,
        userMessage,
        tier: 'medium',
        maxTokens: 512,
        responseFormat: 'text',
        reasoningEffort: 'low',
      });
      return answer.trim();
    } catch (error) {
      this.logger.error('SmartDataQuery LLM error', error);
      return 'I was unable to compute the answer from the sheet data. Please try again.';
    }
  }

  /**
   * Deterministic, no-LLM answer built directly from the sheet list already
   * in `workbookContext`. TASKS.md #257 — `isHidden` used to be dropped
   * before it ever reached the backend (read correctly client-side via
   * Office.js's `worksheet.visibility`, then silently lost converting into
   * this lighter-weight shape), so a hidden sheet's name would appear in the
   * list with no way to say it was hidden — a live-tested false impression
   * that a 6-sheet workbook only had "4 sheets" the user could see, when 2
   * of the 6 were actually hidden ones sitting in the same flat list.
   * `isHidden` is optional — undefined means "unknown" (an older/minimal
   * context that never populated it), counted with the visible sheets rather
   * than guessed as hidden, since most sheets are visible and a false
   * "hidden" claim is worse than an honest gap.
   */
  private answerSheetCountQuestion(workbookContext: WorkbookContext | undefined): string | null {
    const sheets = (workbookContext?.sheets ?? []).filter((s) => s.sheetName);
    if (sheets.length === 0) return null;

    const hidden = sheets.filter((s) => s.isHidden === true);
    const visible = sheets.filter((s) => s.isHidden !== true);
    const quoteAll = (list: typeof sheets) => list.map((s) => `"${s.sheetName}"`).join(', ');

    if (hidden.length === 0) {
      return `This workbook has ${sheets.length} sheet${sheets.length === 1 ? '' : 's'}: ${quoteAll(sheets)}. No hidden sheets.`;
    }

    return (
      `This workbook has ${sheets.length} sheets total — ${visible.length} visible: ${quoteAll(visible)}; ` +
      `${hidden.length} hidden: ${quoteAll(hidden)}.`
    );
  }
}
