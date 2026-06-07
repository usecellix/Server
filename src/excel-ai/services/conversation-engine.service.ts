import { Injectable } from '@nestjs/common';
import { AppConfigService } from '../../config/app-config.service';
import { LlmRequestError } from '../errors/llm-request.error';
import { buildActionPreviewPrompt, buildCellixSystemPrompt } from '../prompt/cellix-system-prompt';
import { extractJsonFromLlmText, hasActionPayload } from '../utils/parse-llm-response.util';
import {
  actionsNeedTableFallback,
  buildTableActionsFromMessage,
  parseTableCreateRequest,
} from '../utils/table-request.util';
import { ConversationMessageEntry } from '../schemas/conversation.schema';
import { SheetActionPayload } from '../types/sheet-actions.types';
import { buildWorkbookContext } from '../utils/workbook-context.util';
import { formatIndianCurrency } from '../utils/indian-format.util';
import { DataQueryService } from './data-query.service';
import { IntentClassifierService, intentIsReadOnly } from './intent-classifier.service';
import { LlmCallTelemetry, LlmUsage, OpenRouterChatMessage, OpenRouterService } from './openrouter.service';
import { SheetAnalysis, SheetAnalyzerService } from './sheet-analyzer.service';

export { LlmRequestError, LlmRequestError as OpenAiRequestError } from '../errors/llm-request.error';
export type { SheetActionPayload };

export type EngineResponse =
  | { kind: 'question'; question: string; options?: string[]; pendingIntent?: string }
  | { kind: 'answer'; answer: string; followUp?: string }
  | { kind: 'actions'; answer: string; actions: SheetActionPayload[]; explanation: string };

type LlmModelTier = 'low' | 'medium' | 'high';

export interface WorkbookContextInput {
  activeSheet?: string;
  sheets?: string[];
}

@Injectable()
export class ConversationEngineService {
  constructor(
    private readonly sheetAnalyzer: SheetAnalyzerService,
    private readonly config: AppConfigService,
    private readonly openRouter: OpenRouterService,
    private readonly intentClassifier: IntentClassifierService,
    private readonly dataQuery: DataQueryService,
  ) {}

  hasOpenAi(): boolean {
    return this.config.hasLlmProvider;
  }

  decide(
    message: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    history: ConversationMessageEntry[],
    workbookMeta?: WorkbookContextInput,
  ): EngineResponse {
    const normalized = message.trim();
    const lower = normalized.toLowerCase();
    const ctx = buildWorkbookContext(sheetData, analysis, workbookMeta);
    const lastAssistant = [...history].reverse().find((entry) => entry.role === 'assistant');

    if (lastAssistant?.metadata?.pendingIntent === 'sum_column') {
      return this.handlePendingSumColumn(normalized, sheetData, analysis);
    }

    const classification = this.intentClassifier.classify(normalized);

    if (intentIsReadOnly(classification.intent)) {
      if (classification.intent === 'EXPLAIN') {
        return { kind: 'answer', answer: this.buildSheetExplanation(sheetData, analysis, ctx) };
      }
      if (classification.intent === 'DATA_QUESTION') {
        const result = this.dataQuery.query(classification.subIntent, normalized, sheetData, analysis, ctx);
        if (result) {
          return { kind: 'answer', answer: result.answer, followUp: result.followUp };
        }
      }
    }

    if (classification.intent === 'FIX') {
      return this.buildFixResponse(normalized, sheetData, analysis, classification.subIntent);
    }

    const cellMatch = /\b([A-Za-z]+\d+)\b/.exec(normalized);
    if (cellMatch && (lower.includes('what') || lower.includes('cell') || lower.includes('value'))) {
      const value = this.sheetAnalyzer.getCellValue(sheetData, cellMatch[1]);
      const display = value === undefined || value === null || value === '' ? '(empty)' : String(value);
      return {
        kind: 'answer',
        answer: `Cell ${cellMatch[1].toUpperCase()} contains "${display}".`,
      };
    }

    if (this.isSumIntent(lower)) {
      return this.handleSumIntent(normalized, lower, sheetData, analysis);
    }

    if (lower.includes('sheet') && (lower.includes('list') || lower.includes('what'))) {
      const sheetList = ctx.sheets.map((s) => `• ${s}`).join('\n');
      return {
        kind: 'answer',
        answer: `Sheets in this workbook:\n${sheetList}\n\nActive sheet: **${ctx.activeSheet}**`,
      };
    }

    if (lower.includes('sort')) {
      return {
        kind: 'question',
        question: 'Sort by which column?',
        options: analysis.headers.slice(0, Math.min(analysis.columnCount, 6)),
        pendingIntent: 'sort_column',
      };
    }

    if (this.isAddRowIntent(lower)) {
      return this.buildAddRowDecision(normalized, lower, analysis);
    }

    if (lower.includes('blank row') && this.isWriteIntent(lower)) {
      return this.buildDeleteBlankRowsDecision(sheetData, analysis);
    }

    const tableActions = buildTableActionsFromMessage(normalized);
    if (tableActions?.length) {
      const plan = parseTableCreateRequest(normalized);
      return {
        kind: 'actions',
        answer: plan
          ? `I'll create **${plan.rowCount}** rows with columns: ${plan.headers.join(', ')}.`
          : "I'll create your table with headers and sample data.",
        explanation: 'Write headers and data rows to the sheet.',
        actions: tableActions,
      };
    }

    if (analysis.isEmpty && this.isPopulateIntent(lower)) {
      return {
        kind: 'question',
        question:
          'I can generate data, but the AI service is not available. Set OPENROUTER_API_KEY in the backend .env and restart the server.',
        options: ['Retry after configuring API key'],
      };
    }

    if (analysis.isEmpty) {
      return {
        kind: 'answer',
        answer:
          'Your worksheet is empty. Ask me to **generate sample GST data**, **create a purchase register**, or describe what columns and rows you need — I will build it for you.',
      };
    }

    const previewHeaders = analysis.headers.filter(Boolean).slice(0, 5).join(', ');
    return {
      kind: 'answer',
      answer: `I see ${analysis.rowCount} rows and ${analysis.columnCount} columns${previewHeaders ? ` (${previewHeaders})` : ''}. Ask about a cell, a column total, or describe a change you want.`,
    };
  }

  async *streamOpenAiAnswer(
    message: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    history: ConversationMessageEntry[],
    telemetry?: LlmCallTelemetry,
    workbookMeta?: WorkbookContextInput,
  ): AsyncGenerator<string> {
    const messages = this.buildLlmMessages(message, sheetData, analysis, history, workbookMeta);
    const tier = this.selectModelTier(message, analysis, history);

    if (this.openRouter.isConfigured()) {
      if (telemetry) telemetry.modelTier = tier;
      const maxTokens = this.selectMaxTokens(message, analysis);
      yield* this.openRouter.streamChat(
        messages,
        telemetry,
        this.getOpenRouterModelForTier(tier),
        maxTokens,
      );
      return;
    }

    yield* this.streamOpenAiDirect(messages, telemetry, tier);
  }

  private buildLlmMessages(
    message: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    history: ConversationMessageEntry[],
    workbookMeta?: WorkbookContextInput,
  ): OpenRouterChatMessage[] {
    const ctx = buildWorkbookContext(sheetData, analysis, workbookMeta);
    const classification = this.intentClassifier.classify(message);
    const systemPrompt = `${buildCellixSystemPrompt(ctx, analysis.isEmpty)}

${buildActionPreviewPrompt(classification.intent)}

Sheet is ${analysis.isEmpty ? 'EMPTY — populate with SET_CELL row 0 for headers and ADD_ROW for data' : 'not empty'}.
Sheet has ${analysis.rowCount} rows, ${analysis.columnCount} columns. Next append row index: ${Math.max(analysis.rowCount, 1)}.`;

    const prior = history.slice(-12).map((entry) => ({
      role: entry.role as 'user' | 'assistant',
      content: entry.content,
    }));

    return [
      { role: 'system', content: systemPrompt },
      ...prior,
      {
        role: 'user',
        content: `User message: ${message}\nSheet preview (first rows): ${JSON.stringify(sheetData.slice(0, 8))}`,
      },
    ];
  }

  parseStructuredResponse(
    text: string,
    analysis?: SheetAnalysis,
    userMessage?: string,
  ): EngineResponse | null {
    const parsed = extractJsonFromLlmText(text);
    if (!parsed && userMessage) {
      return this.tableFallbackResponse(userMessage, text);
    }
    if (!parsed) return null;

    if (parsed.type === 'question' && parsed.question) {
      return { kind: 'question', question: parsed.question, options: parsed.options };
    }

    const actionList = hasActionPayload(parsed) ? parsed.actions! : null;
    const isActionsType =
      parsed.type === 'actions' || (actionList !== null && parsed.type !== 'answer');

    if (isActionsType && actionList) {
      let finalActions = actionList;
      if (userMessage && actionsNeedTableFallback(userMessage, actionList)) {
        const fallback = buildTableActionsFromMessage(userMessage);
        if (fallback) finalActions = fallback;
      }

      const sanitized = this.sanitizeActions(finalActions, analysis);
      if (sanitized.length === 0) {
        const emptyHint = analysis?.isEmpty
          ? 'I could not apply those changes to an empty sheet. Try asking again: "Generate 10 rows of sample GST purchase data with headers".'
          : 'Where should the new row go, and what values should it contain?';
        return {
          kind: 'question',
          question: emptyHint,
          options: analysis?.isEmpty
            ? ['Generate 10 rows of GST sample data', 'Generate 5 rows with headers only']
            : ['After the last row (dummy values)', 'After the last row (I will specify values)'],
        };
      }
      return {
        kind: 'actions',
        answer: parsed.answer || parsed.explanation || 'Here are the changes I will apply to your sheet.',
        actions: sanitized,
        explanation: parsed.explanation || parsed.answer || 'Applied sheet changes.',
      };
    }

    if (parsed.answer && !actionList) {
      if (userMessage) {
        const tableOnly = this.tableFallbackResponse(userMessage, parsed.answer);
        if (tableOnly?.kind === 'actions') return tableOnly;
      }
      return { kind: 'answer', answer: parsed.answer };
    }

    if (userMessage) {
      return this.tableFallbackResponse(userMessage, text);
    }

    return null;
  }

  private tableFallbackResponse(message: string, hint: string): EngineResponse | null {
    const tableActions = buildTableActionsFromMessage(message);
    if (!tableActions) return null;

    const plan = parseTableCreateRequest(message);
    const intro = plan
      ? `Created **${plan.rowCount}** rows with columns: ${plan.headers.join(', ')}.`
      : 'Created your table with headers and sample values.';

    return {
      kind: 'actions',
      answer: hint.length > 10 && !hint.startsWith('{') ? `${hint}\n\n${intro}` : intro,
      explanation: 'Write headers and all data rows in one step.',
      actions: tableActions,
    };
  }

  private static readonly HEADER_ROW = 0;

  private sanitizeActions(
    actions: SheetActionPayload[],
    analysis?: SheetAnalysis,
  ): SheetActionPayload[] {
    const withAddRowConversion = this.convertHeaderRowWritesToAddRow(actions, analysis);
    return withAddRowConversion
      .map((action) => this.sanitizeAction(action))
      .filter((action): action is SheetActionPayload => action !== null)
      .filter((action) => !this.isHeaderMutation(action, analysis?.isEmpty));
  }

  private isHeaderMutation(action: SheetActionPayload, sheetIsEmpty = false): boolean {
    if (sheetIsEmpty) {
      const allowedOnHeader = new Set(['SET_CELL', 'SET_FORMULA', 'FORMAT_RANGE', 'MERGE_CELLS']);
      if (action.row === ConversationEngineService.HEADER_ROW && allowedOnHeader.has(action.type)) {
        return false;
      }
    }

    const rowOnlyTypes = new Set([
      'ADD_ROW',
      'CREATE_SHEET',
      'DELETE_SHEET',
      'RENAME_SHEET',
      'COPY_SHEET',
      'HIDE_SHEET',
      'SHOW_SHEET',
      'SET_SHEET_COLOR',
      'UNFREEZE_PANES',
      'FREEZE_PANES',
    ]);
    if (rowOnlyTypes.has(action.type)) return false;
    if (action.type === 'WRITE_TABLE') return false;
    if (action.type === 'MERGE_CELLS' || action.type === 'FORMAT_RANGE') return false;
    return action.row === ConversationEngineService.HEADER_ROW;
  }

  private convertHeaderRowWritesToAddRow(
    actions: SheetActionPayload[],
    analysis?: SheetAnalysis,
  ): SheetActionPayload[] {
    if (analysis?.isEmpty) {
      return actions;
    }

    const headerWrites = actions.filter(
      (action) =>
        (action.type === 'SET_CELL' ||
          action.type === 'SET_FORMULA' ||
          action.type === 'CLEAR_CELL') &&
        action.row === ConversationEngineService.HEADER_ROW,
    );

    if (!headerWrites.length) return actions;

    const columnCount = Math.max(
      analysis?.columnCount ?? 0,
      ...headerWrites.map((action) => (action.col ?? 0) + 1),
      1,
    );
    const rowData: unknown[] = Array.from({ length: columnCount }, (_, index) => {
      const write = headerWrites.find((action) => action.col === index);
      if (!write) return '';
      if (write.type === 'SET_FORMULA') return write.formula ?? '';
      if (write.type === 'SET_CELL') return write.value ?? '';
      return '';
    });

    const rest = actions.filter((action) => !headerWrites.includes(action));
    return [{ type: 'ADD_ROW', data: rowData }, ...rest];
  }

  private sanitizeAction(action: SheetActionPayload): SheetActionPayload | null {
    const row = this.normalizeIndex(action.row);
    const col = this.normalizeIndex(action.col);

    switch (action.type) {
      case 'SET_CELL':
        if (row === undefined || col === undefined || action.value === undefined) return null;
        return { ...action, row, col };
      case 'CLEAR_CELL':
        if (row === undefined || col === undefined) return null;
        return { ...action, row, col };
      case 'SET_FORMULA':
        if (row === undefined || col === undefined || typeof action.formula !== 'string') return null;
        return { ...action, row, col };
      case 'ADD_ROW':
        if (!Array.isArray(action.data)) return null;
        return action;
      case 'DELETE_ROW':
      case 'INSERT_ROW':
      case 'HIDE_ROW':
      case 'SHOW_ROW':
      case 'SET_ROW_HEIGHT':
        if (row === undefined) return null;
        return { ...action, row };
      case 'INSERT_COLUMN':
      case 'DELETE_COLUMN':
      case 'HIDE_COLUMN':
      case 'SHOW_COLUMN':
      case 'SET_COLUMN_WIDTH':
      case 'FILL_DOWN':
        if (col === undefined) return null;
        return { ...action, col };
      case 'FILL_RIGHT':
        if (row === undefined || col === undefined) return null;
        return { ...action, row, col };
      case 'HIGHLIGHT_CELL':
        if (row === undefined || col === undefined) return null;
        return { ...action, row, col, color: action.color || '#DCFCE7' };
      case 'FORMAT_RANGE':
      case 'MERGE_CELLS':
      case 'UNMERGE_CELLS':
      case 'CLEAR_CONTENT':
      case 'CLEAR_FORMAT':
      case 'CLEAR_ALL':
      case 'ADD_COMMENT':
      case 'DELETE_COMMENT':
        if (row === undefined || col === undefined) return null;
        return { ...action, row, col };
      case 'WRITE_TABLE':
        if (!Array.isArray(action.headers) || !Array.isArray(action.rows)) return null;
        return action;
      case 'FREEZE_PANES':
      case 'UNFREEZE_PANES':
      case 'CREATE_SHEET':
      case 'DELETE_SHEET':
      case 'RENAME_SHEET':
      case 'COPY_SHEET':
      case 'HIDE_SHEET':
      case 'SHOW_SHEET':
      case 'SET_SHEET_COLOR':
        return action;
      default:
        return null;
    }
  }

  private normalizeIndex(value: unknown): number | undefined {
    if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) {
      return undefined;
    }
    return value;
  }

  private handlePendingSumColumn(
    normalized: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
  ): EngineResponse {
    const columnIndex = this.sheetAnalyzer.resolveColumnIndex(normalized, analysis);
    if (columnIndex === null) {
      return {
        kind: 'question',
        question: 'I did not recognize that column. Which column should I total?',
        options: analysis.columnLetters.slice(0, Math.min(analysis.columnCount, 6)),
        pendingIntent: 'sum_column',
      };
    }

    const total = this.sheetAnalyzer.sumColumn(sheetData, columnIndex);
    const columnLabel = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];
    const totalRow = this.buildSummaryRow(analysis, columnIndex, total, `Total ${columnLabel}`);
    return {
      kind: 'actions',
      answer: `The total for ${columnLabel} is ${formatIndianCurrency(total)}. I'll add a summary row.`,
      explanation: `Add a new summary row with the total for ${columnLabel}.`,
      actions: [{ type: 'ADD_ROW', data: totalRow }],
    };
  }

  private handleSumIntent(
    normalized: string,
    lower: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
  ): EngineResponse {
    if (analysis.isEmpty || analysis.columnCount === 0) {
      return {
        kind: 'answer',
        answer: 'Your sheet looks empty. Add data with numeric values, then ask me to calculate a total.',
      };
    }

    const explicitColumn = this.extractColumnReference(normalized, analysis);
    if (explicitColumn !== null) {
      const total = this.sheetAnalyzer.sumColumn(sheetData, explicitColumn);
      const columnLabel = analysis.headers[explicitColumn] || analysis.columnLetters[explicitColumn];
      if (this.isWriteResultIntent(lower)) {
        return {
          kind: 'actions',
          answer: `The total for ${columnLabel} is ${formatIndianCurrency(total)}.`,
          explanation: `Add a new summary row with the total for ${columnLabel}.`,
          actions: [
            {
              type: 'ADD_ROW',
              data: this.buildSummaryRow(analysis, explicitColumn, total, `Total ${columnLabel}`),
            },
          ],
        };
      }
      return {
        kind: 'answer',
        answer: `The total for ${columnLabel} is ${formatIndianCurrency(total)}.`,
        followUp: `Want me to add a SUM formula at the bottom of column ${analysis.columnLetters[explicitColumn]}?`,
      };
    }

    return {
      kind: 'question',
      question: 'Which column contains the values you want to total?',
      options: analysis.headers.slice(0, Math.min(analysis.columnCount, 6)),
      pendingIntent: 'sum_column',
    };
  }

  private buildFixResponse(
    message: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    fixType?: string,
  ): EngineResponse {
    const columnIndex = this.sheetAnalyzer.resolveColumnIndexFromMessage(message, analysis);

    if (fixType === 'text_numbers' && columnIndex !== null) {
      const letter = analysis.columnLetters[columnIndex];
      const label = analysis.headers[columnIndex] || letter;
      const startRow = analysis.rowCount > 1 ? 2 : 1;
      const endRow = analysis.rowCount;
      return {
        kind: 'actions',
        answer: `Column **${label}** likely has text-stored numbers (common in Tally exports). SUM returns 0 because Excel treats them as text.\n\nHere's the fix:\n• Range: ${letter}${startRow}:${letter}${endRow}\n• Formula: =VALUE(SUBSTITUTE(${letter}2,"₹",""))`,
        explanation: `Convert text-stored numbers in column ${letter} to actual numbers using VALUE(SUBSTITUTE()).`,
        actions: [
          {
            type: 'SET_FORMULA',
            row: 1,
            col: columnIndex,
            formula: `=VALUE(SUBSTITUTE(${letter}2,"₹",""))`,
          },
          {
            type: 'FILL_DOWN',
            col: columnIndex,
            row: 1,
            endRow: Math.max(analysis.rowCount - 1, 1),
          },
        ],
      };
    }

    if (fixType === 'text_dates' && columnIndex !== null) {
      const letter = analysis.columnLetters[columnIndex];
      const label = analysis.headers[columnIndex] || letter;
      return {
        kind: 'actions',
        answer: `Dates in **${label}** are likely stored as text (common in Tally exports), which breaks sorting.\n\nFix: convert with DATEVALUE().`,
        explanation: `Convert text dates in column ${letter} to real Excel dates.`,
        actions: [
          {
            type: 'SET_FORMULA',
            row: 1,
            col: columnIndex,
            formula: `=DATEVALUE(${letter}2)`,
          },
          {
            type: 'FILL_DOWN',
            col: columnIndex,
            row: 1,
            endRow: Math.max(analysis.rowCount - 1, 1),
          },
        ],
      };
    }

    if (fixType === 'na_error') {
      return {
        kind: 'answer',
        answer:
          '**#N/A** usually means the lookup value was not found. Common causes:\n• Trailing spaces — wrap lookup value in TRIM()\n• Type mismatch — wrap in VALUE() if comparing numbers\n• Fix: wrap your formula in IFERROR(formula,"Not Found")',
        followUp: 'Tell me which cell has the error and I can propose a corrected formula.',
      };
    }

    if (fixType === 'ref_error') {
      return {
        kind: 'answer',
        answer:
          '**#REF!** means a formula references a deleted row or column. Check recent insert/delete operations and update the formula range.',
        followUp: 'Tell me which column has the error and I can trace the broken reference.',
      };
    }

    return {
      kind: 'answer',
      answer:
        'Describe the error or the column/cell that is broken, and I will diagnose and propose a fix.',
    };
  }

  private buildDeleteBlankRowsDecision(
    sheetData: unknown[][],
    analysis: SheetAnalysis,
  ): EngineResponse {
    const blankRows = this.sheetAnalyzer.findBlankRows(sheetData);
    if (blankRows.length === 0) {
      return { kind: 'answer', answer: 'No completely blank rows found in your data.' };
    }

    const preview = blankRows
      .slice(0, 5)
      .map((r) => `row ${r + 1}`)
      .join(', ');
    const more = blankRows.length > 5 ? ` and ${blankRows.length - 5} more` : '';

    return {
      kind: 'actions',
      answer: `Found **${blankRows.length}** blank row(s): ${preview}${more}.\n\nApprove to delete all ${blankRows.length} blank rows?`,
      explanation: `Delete ${blankRows.length} blank data row(s).`,
      actions: blankRows.map((row) => ({ type: 'DELETE_ROW' as const, row })),
    };
  }

  private buildSummaryRow(
    analysis: SheetAnalysis,
    valueColumnIndex: number,
    value: unknown,
    label: string,
  ): unknown[] {
    const columnCount = Math.max(analysis.columnCount, valueColumnIndex + 1, 1);
    const row: unknown[] = Array.from({ length: columnCount }, () => '');
    const labelColumn = valueColumnIndex > 0 ? 0 : Math.min(1, columnCount - 1);
    row[labelColumn] = label;
    row[valueColumnIndex] = value;
    return row;
  }

  private buildSheetExplanation(
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    ctx: ReturnType<typeof buildWorkbookContext>,
  ): string {
    if (analysis.isEmpty) {
      return 'Your worksheet appears empty. Add headers and data, then ask me to explain or analyze it again.';
    }

    const MAX_LISTED_COLUMNS = 10;
    const MAX_SAMPLES = 2;
    const MAX_VALUE_CHARS = 22;

    const populatedRows = sheetData.filter(
      (row) =>
        Array.isArray(row) &&
        row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== ''),
    ).length;

    const clip = (s: string): string => {
      const t = s.trim();
      if (t.length <= MAX_VALUE_CHARS) return t;
      return `${t.slice(0, MAX_VALUE_CHARS - 1)}…`;
    };

    type ColEntry = { header: string; letter: string; preview: string };
    const columnsWithData: ColEntry[] = [];

    analysis.headers.forEach((header, columnIndex) => {
      const samples: string[] = [];
      for (let row = 0; row < sheetData.length && samples.length < MAX_SAMPLES; row += 1) {
        const value = sheetData[row]?.[columnIndex];
        if (value === null || value === undefined || String(value).trim() === '') {
          continue;
        }
        samples.push(clip(String(value).trim()));
      }
      if (samples.length === 0) {
        return;
      }
      const letter = analysis.columnLetters[columnIndex] ?? String(columnIndex + 1);
      columnsWithData.push({
        header,
        letter,
        preview: samples.join(' · '),
      });
    });

    const detailed = columnsWithData.slice(0, MAX_LISTED_COLUMNS);
    const overflowCols = columnsWithData.slice(MAX_LISTED_COLUMNS);

    const safe = (s: string) => s.replace(/\*\*/g, '');

    const blocks: string[] = [];

    blocks.push(
      `**${ctx.activeSheet}** — **${analysis.rowCount}** rows, **${analysis.columnCount}** columns; **${populatedRows}** row(s) with data. Range: ${ctx.dataRange}`,
    );

    if (detailed.length === 0) {
      blocks.push(
        'Most cells look **empty** or free-form — add headers and structured rows to get column-level insight.',
      );
    } else {
      const lines = detailed.map((c) => {
        const h = safe(c.header);
        return `• **${h}** (${c.letter}): ${c.preview}`;
      });
      blocks.push(
        `**${detailed.length}** of **${columnsWithData.length}** columns show sample values:\n${lines.join('\n')}`,
      );
      if (overflowCols.length > 0) {
        const names = overflowCols
          .slice(0, 8)
          .map((c) => `**${safe(c.header)}**`)
          .join(', ');
        const more =
          overflowCols.length > 8 ? ` (+**${overflowCols.length - 8}** more)` : '';
        blocks.push(`Also present: ${names}${more}.`);
      }
    }

    blocks.push(
      'Try **"What is in B3?"**, a **column total**, or describe a **change** you want.',
    );

    return blocks.join('\n\n');
  }

  private isSumIntent(lower: string): boolean {
    return (
      lower.includes('total') ||
      lower.includes('sum') ||
      lower.includes('add up') ||
      (lower.includes('calculate') && !lower.includes('formula'))
    );
  }

  private isWriteResultIntent(lower: string): boolean {
    return /\b(create|write|put|add|insert|show|place|make|set)\b/.test(lower);
  }

  private isWriteIntent(lower: string): boolean {
    return /\b(add|create|insert|delete|remove|rename|copy|move|hide|unhide|show|format|apply|make|set|put|write|change|update|merge|unmerge|clear|sort|filter|freeze|paste|fill|highlight|lock|unlock|protect|resize|wrap)\b/.test(
      lower,
    );
  }

  private isAddRowIntent(lower: string): boolean {
    return (
      /\b(add|create|insert|append|new)\b.*\b(row|record|entry|line)\b/.test(lower) ||
      /\b(row|record|entry)\b.*\b(dummy|sample|placeholder|test|example)\b/.test(lower) ||
      /\b(dummy|sample|placeholder|test)\b.*\b(row|values?|data)\b/.test(lower)
    );
  }

  private buildAddRowDecision(
    _message: string,
    lower: string,
    analysis: SheetAnalysis,
  ): EngineResponse {
    if (analysis.isEmpty || analysis.columnCount === 0) {
      return {
        kind: 'question',
        question: 'Your sheet looks empty. Should I add a header row first, or append a row to a specific location?',
        options: ['Add headers and one dummy data row', 'Tell me where to add the row'],
      };
    }

    const wantsDummy =
      /\b(dummy|sample|placeholder|test|example|fake)\b/.test(lower) ||
      !/\b(with|value|values|containing|for)\b/.test(lower);

    if (!wantsDummy) {
      return {
        kind: 'question',
        question: 'What values should the new row contain? I will add it after your existing rows without changing the header.',
        options: [
          'Use placeholder dummy values',
          'After the last row — I will type the values next',
        ],
        pendingIntent: 'add_row',
      };
    }

    return {
      kind: 'actions',
      answer:
        'I prepared a new data row with placeholder values. It will be added after your existing rows — your header row will stay unchanged.',
      explanation: 'Append a new data row with dummy values aligned to your column headers.',
      actions: [{ type: 'ADD_ROW', data: this.buildDummyRow(analysis) }],
    };
  }

  private buildDummyRow(analysis: SheetAnalysis): unknown[] {
    const columnCount = Math.max(analysis.columnCount, 1);
    return Array.from({ length: columnCount }, (_, index) => {
      const header = analysis.headers[index] || analysis.columnLetters[index] || `Column ${index + 1}`;
      const label = String(header).replace(/^Column\s+/i, '').trim() || `Col${index + 1}`;
      return `Sample ${label}`;
    });
  }

  private extractColumnReference(message: string, analysis: SheetAnalysis): number | null {
    return this.sheetAnalyzer.resolveColumnIndexFromMessage(message, analysis);
  }

  private async *streamOpenAiDirect(
    messages: OpenRouterChatMessage[],
    telemetry?: LlmCallTelemetry,
    tier: LlmModelTier = 'medium',
  ): AsyncGenerator<string> {
    const apiKey = this.config.openAiApiKey;
    if (!apiKey) return;
    const model = this.getOpenAiModelForTier(tier);

    if (telemetry) {
      telemetry.provider = 'openai';
      telemetry.model = model;
      telemetry.modelTier = tier;
    }

    let response: Response;
    try {
      response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${apiKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model,
          messages,
          stream: true,
          stream_options: { include_usage: true },
          temperature: 0.25,
          max_tokens: 4096,
        }),
      });
    } catch {
      throw new LlmRequestError(503, 'OpenAI network error');
    }

    if (!response.ok || !response.body) {
      throw new LlmRequestError(response.status);
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        if (!line.startsWith('data: ')) continue;
        const payload = line.slice(6).trim();
        if (payload === '[DONE]') return;

        try {
          const parsed = JSON.parse(payload) as {
            choices?: Array<{ delta?: { content?: string } }>;
            usage?: {
              prompt_tokens?: number;
              completion_tokens?: number;
              total_tokens?: number;
            } | null;
          };
          if (parsed.usage && telemetry) {
            telemetry.usage = this.normalizeOpenAiUsage(parsed.usage);
          }
          const token = parsed.choices?.[0]?.delta?.content;
          if (token) yield token;
        } catch {
          // ignore malformed chunks
        }
      }
    }
  }

  private normalizeOpenAiUsage(usage: {
    prompt_tokens?: number;
    completion_tokens?: number;
    total_tokens?: number;
  }): LlmUsage {
    return {
      promptTokens: usage.prompt_tokens,
      completionTokens: usage.completion_tokens,
      totalTokens: usage.total_tokens,
    };
  }

  private selectModelTier(
    message: string,
    analysis: SheetAnalysis,
    history: ConversationMessageEntry[],
  ): LlmModelTier {
    const lower = message.toLowerCase();
    const editIntent =
      /\b(set|change|update|edit|delete|clear|remove|create|add|insert|formula|highlight|fill|format|merge|sort|filter)\b/.test(
        lower,
      );
    const complexIntent =
      /\b(reconcile|audit|compare|match|variance|all rows|entire sheet|bulk|every|generate report|analy[sz]e deeply|pivot|vlookup|sumif|countif)\b/.test(
        lower,
      );

    if (complexIntent || analysis.rowCount > 200 || history.length > 20) {
      return 'high';
    }

    if (editIntent || analysis.rowCount > 50 || history.length > 8) {
      return 'medium';
    }

    return 'low';
  }

  private getOpenRouterModelForTier(tier: LlmModelTier): string {
    const model =
      tier === 'low'
        ? this.config.openRouterModelLow
        : tier === 'high'
          ? this.config.openRouterModelHigh
          : this.config.openRouterModelMedium;
    if (model.includes('openrouter/auto')) {
      return tier === 'high' ? 'openai/gpt-5' : 'openai/gpt-5-mini';
    }
    return model;
  }

  private getOpenAiModelForTier(tier: LlmModelTier): string {
    if (tier === 'low') return this.config.openAiModelLow;
    if (tier === 'high') return this.config.openAiModelHigh;
    return this.config.openAiModelMedium;
  }

  private selectMaxTokens(message: string, analysis: SheetAnalysis): number {
    const lower = message.toLowerCase();
    const isGenerative =
      analysis.isEmpty ||
      /\b(generate|populate|dummy|sample|random|seed|mock|create.*data|fill.*sheet)\b/.test(lower);
    return isGenerative ? 8192 : 4096;
  }

  private isPopulateIntent(lower: string): boolean {
    return /\b(generate|populate|dummy|sample|random|seed|mock|create.*data|gst)\b/.test(lower);
  }
}
