import { Injectable, Logger } from '@nestjs/common';
import { WorkbookContext } from '../../types/cellix.types';
import {
  buildDataQuerySystemPrompt,
  buildDataQueryUserMessage,
} from '../prompts/data-query-system-prompt';
import { sliceRelevantColumns } from '../utils/column-slicer.util';
import { isFindLookupMessage } from '../utils/find-query-parser.util';
import { ColumnHint, isDateNumberFormat, serialToIso } from '../table-query/table-model';
import { formatDate } from '../table-query/table-query.format';
import { looksLikeTableQuestion, TableQuestionAnswerer } from '../table-query/table-question';
import { FindMatch } from './data-query.service';
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

/** Shown instead of a blank when the model returns no text. TASKS.md #382. */
export const NO_ANSWER_MESSAGE =
  'I could not work out an answer to that from the sheet. Try asking for one figure at a time, for example "What is the total of the Debit column?"';

const norm = (value: unknown): string => String(value ?? '').trim().toLowerCase();

/** A computed answer and the rows it names, as pointers the add-in can jump to. */
export interface ComputedTableAnswer {
  answer: string;
  matches: FindMatch[];
}

const MAX_POINTERS = 60;

function columnLetter(index: number): string {
  let n = index + 1;
  let letters = '';
  while (n > 0) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/** True when the sheet's used range starts at A1, so a row of the data is the same row on the sheet. */
function startsAtA1(usedRange: string | undefined): boolean {
  if (!usedRange) return true;
  const start = usedRange.replace(/^.*!/, '').replace(/\$/g, '').split(':')[0].toUpperCase();
  return start === 'A1' || start === '';
}

/** The row of `sheetData` that holds the column headings. */
function findHeaderRow(sheetData: unknown[][], headers: string[] | undefined, hint: number | undefined): number {
  const wanted = (headers ?? []).map(norm).filter(Boolean);
  if (!wanted.length) return hint ?? 0;
  const matchesAt = (index: number) => {
    const row = (sheetData[index] ?? []).map(norm);
    return wanted.filter((h) => row.includes(h)).length / wanted.length >= 0.8;
  };
  if (hint !== undefined && matchesAt(hint)) return hint;
  for (let index = 0; index < Math.min(sheetData.length, 30); index++) if (matchesAt(index)) return index;
  return hint ?? 0;
}

@Injectable()
export class SmartDataQueryService {
  private readonly logger = new Logger(SmartDataQueryService.name);
  private readonly tableQuestions: TableQuestionAnswerer;

  constructor(private readonly openRouter: OpenRouterService) {
    this.tableQuestions = new TableQuestionAnswerer(openRouter);
  }

  /**
   * Answers a question about a sheet's rows by planning a query and running it
   * in code over every row. Returns null when the message is not that kind of
   * question or no plan could be made, so the caller can use its usual path.
   */
  async tryTableAnswer(
    message: string,
    sheetData: unknown[][],
    workbookContext: WorkbookContext | undefined,
    sheetName: string | undefined,
  ): Promise<ComputedTableAnswer | null> {
    if (!sheetData?.length || isFindLookupMessage(message) || !looksLikeTableQuestion(message)) return null;
    const snapshot =
      workbookContext?.sheets?.find((s) => s.sheetName === sheetName) ??
      (workbookContext?.sheets?.length === 1 ? workbookContext.sheets[0] : undefined);
    const columnHints: ColumnHint[] = (snapshot?.columnMeta ?? []).map((meta) => ({
      index: meta.index,
      detectedType: meta.detectedType ? String(meta.detectedType) : undefined,
      numberFormat: meta.numberFormat,
    }));
    const result = await this.tableQuestions.tryAnswer({
      message,
      sheetName: sheetName || snapshot?.sheetName || 'the sheet',
      sheetData,
      headers: snapshot?.headers,
      headerRowIndex: findHeaderRow(sheetData, snapshot?.headers, snapshot?.headerRowIndex),
      columnHints,
      declaredRowCount: snapshot?.rowCount,
    });
    if (!result) return null;

    // Pointers need a sheet to open and rows that line up with it. Without
    // both, the answer still stands; it just has nothing to click.
    const pointerSheet = sheetName || snapshot?.sheetName;
    const matches: FindMatch[] =
      pointerSheet && startsAtA1(snapshot?.usedRange)
        ? result.rowNumbers.slice(0, MAX_POINTERS).map((rowNum) => ({
            label: `row ${rowNum}`,
            sheetName: pointerSheet,
            row: rowNum - 1,
            col: 0,
            colLetter: 'A',
            rowNum,
            rawValue: '',
            endCol: Math.max(0, result.columnCount - 1),
            detail: `A${rowNum}:${columnLetter(Math.max(0, result.columnCount - 1))}${rowNum}`,
          }))
        : [];
    return { answer: result.answer, matches };
  }

  /**
   * Answer a read-only data query using column-sliced sheet data and the MEDIUM LLM tier.
   */
  async handleQuery(
    message: string,
    sheetData: unknown[][],
    workbookContext: WorkbookContext | undefined,
    activeSheetName: string | undefined,
    emit: SmartDataQueryEmit,
    /** Out-param: filled with row pointers when the answer was computed in code. */
    out?: { matches?: FindMatch[] },
  ): Promise<string> {
    if (WORKBOOK_SHEET_COUNT_PATTERN.test(message)) {
      const answer = this.answerSheetCountQuestion(workbookContext);
      if (answer) return answer;
    }

    // Computed in code over every row whenever the question allows it. What
    // follows is the older path, where the model reads the rows itself.
    const computed = await this.tryTableAnswer(message, sheetData, workbookContext, activeSheetName);
    if (computed) {
      if (out) out.matches = computed.matches;
      return computed.answer;
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
    this.showDatesAsDates(sheet, workbookContext);

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
      // A reasoning model can spend its whole budget thinking and return no
      // text. That must never reach the user as a blank answer.
      return answer.trim() || NO_ANSWER_MESSAGE;
    } catch (error) {
      this.logger.error('SmartDataQuery LLM error', error);
      return 'I was unable to compute the answer from the sheet data. Please try again.';
    }
  }

  /**
   * Excel hands dates over as serial numbers (46113 for 1 Apr 2026). A model
   * asked to read them converts in its head and gets it wrong, so a column the
   * add-in reported as a date is rewritten as dates before it is shown.
   * TASKS.md #381.
   */
  private showDatesAsDates(
    sheet: { sheetName: string; columnIndices: number[]; rows: string[][] },
    workbookContext: WorkbookContext | undefined,
  ): void {
    const snapshot = workbookContext?.sheets?.find((s) => s.sheetName === sheet.sheetName);
    const dateColumns = new Set(
      (snapshot?.columnMeta ?? [])
        .filter((meta) => meta.detectedType === 'date' || isDateNumberFormat(meta.numberFormat))
        .map((meta) => meta.index),
    );
    if (!dateColumns.size) return;
    sheet.columnIndices.forEach((original, position) => {
      if (!dateColumns.has(original)) return;
      for (const row of sheet.rows) {
        const cell = row[position];
        if (!/^\d{4,6}(\.\d+)?$/.test(String(cell ?? '').trim())) continue;
        const iso = serialToIso(Number(cell));
        if (iso) row[position] = formatDate(iso);
      }
    });
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
