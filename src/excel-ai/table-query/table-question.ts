import { Logger } from '@nestjs/common';
import { LlmCompletionOutcome, OpenRouterService } from '../services/openrouter.service';
import { buildTable, ColumnHint, isBlank, TableModel, toIsoDateValue } from './table-model';
import { executeTableQuery } from './table-query.executor';
import { formatTableAnswer, rowsMentioned } from './table-query.format';
import { parseTableQueryPlan } from './table-query.plan';
import { TableQueryPlanError } from './table-query.types';

/**
 * Answers a question about a sheet's rows without letting the model touch the
 * numbers. The model reads the column list and translates the question into a
 * query plan; code runs the plan over every row and writes the answer.
 *
 * This replaces "hand the rows to the model and ask for the answer", which
 * produced wrong totals, skipped rows, swapped figures and empty replies
 * (TASKS.md #380 to #383). It follows the rule the domain tools already obey:
 * the model plans, code computes.
 */

const WRITE_WORDS =
  /\b(highlight|format|colou?r|sort|delete|remove|add(?!s?\s+up\b)|insert|create|make|build|generate|copy|move|rename|chart|plot|graph|filter|hide|unhide|freeze|merge|fill|replace|convert|clear|apply|fix|clean)\b/i;

const ANALYTIC_WORDS =
  /\b(total|totals|sum|adds? up|how many|how much|count|number of|average|mean|largest|biggest|highest|greatest|maximum|max|most|smallest|lowest|least|minimum|min|top|bottom|first|last|latest|earliest|oldest|newest|between|per (?:day|month|year)|by (?:day|month|year)|each (?:day|month|year)|monthly|yearly|period|date range|which|spent?|spend|paid|received|credited|debited|balance|transactions?|entries|rows)\b/i;

/** A phrase that opens with the measure itself: "Total debits per month", "Top 5 debits". */
const ANALYTIC_OPENER =
  /^(total|totals|sum|count|average|mean|number of|top|bottom|largest|biggest|highest|lowest|smallest|first|last|earliest|latest)\b/i;

/** "...in a new row", "...into column G": the answer is to be put somewhere, which is a change. */
const WRITE_DESTINATION = /\b(?:in|into|to|at|under|below)\s+(?:a\s+|the\s+)?(?:new\s+|next\s+|last\s+|bottom\s+)?(?:row|column|cell|sheet|tab)\b/i;

const SINGLE_CELL = /\bcell\s+\$?[A-Za-z]{1,3}\$?\d+\b/i;
const WORKBOOK_STRUCTURE = /\b(sheets?|tabs?|workbook)\b.*\b(how many|names?|list)\b|\b(how many|what|list|names? of)\b.*\b(sheets?|tabs?)\b/i;

/**
 * True for a message that asks something of the rows of a table and asks for
 * no change. Cheap by design: it only decides whether a plan is worth asking
 * for, and the planner can still decline.
 */
export function looksLikeTableQuestion(message: string): boolean {
  const text = String(message ?? '').trim();
  if (!text || text.length > 600) return false;
  if (SINGLE_CELL.test(text) || WORKBOOK_STRUCTURE.test(text)) return false;
  const asks = /\?\s*$/.test(text) || /^(what|which|how|when|who|whose|where|is|are|was|were|do|does|did|any|list|show me|give me|tell me)\b/i.test(text) || ANALYTIC_OPENER.test(text);
  if (!asks) return false;
  if (WRITE_WORDS.test(text)) return false;
  if (ANALYTIC_OPENER.test(text) && WRITE_DESTINATION.test(text)) return false;
  return ANALYTIC_WORDS.test(text);
}

export const TABLE_QUERY_SYSTEM_PROMPT = `You translate a question about ONE spreadsheet table into a JSON query plan.
You never compute, estimate or quote a value from the data. Code runs your plan over every row.

Reply with JSON only, in one of these two shapes:
{"queries":[ ... ]}            one to six queries, in the order the question asks
{"unsupported":"short reason"}

A query: {"op": "...", "column": "...", "filters": [ ... ], "groupBy": {"column": "...", "by": "..."}, "limit": N}

op:
- sum, average: of a number column.
- count: rows left after the filters. With "column", rows where that column is not blank.
- min, max: lowest or highest value of a number or date column. The row it sits on is returned too, so use these for "largest debit and what was it for", "lowest balance and when", "earliest / latest date".
- top, bottom: the N rows with the largest or smallest values in "column". Set "limit" (default 5).
- list: the rows matching the filters ("limit" default 20).
- first, last: the value of "column" in the first or last row. The closing balance of a statement is the last Balance.
- distinct: the different values of a column.

filters (every one must hold): {"column": "...", "op": "...", "value": ..., "value2": ...}
ops: contains, not_contains, equals, not_equals, starts_with, ends_with, gt, gte, lt, lte, between (value and value2), blank, not_blank.
Text matching ignores case. Write dates as "YYYY-MM-DD". A month is "between" its first and last day.

groupBy: {"column": "...", "by": "value" | "day" | "month" | "year"}. day, month and year are for date columns. Use it with sum, average, count, min or max for "per month", "for each payee" and similar.
The groupBy column is what the answer is split BY, never the column being added up. "Total debits per month" is {"op":"sum","column":"Debit","groupBy":{"column":"Date","by":"month"}}: Debit is added up, the Date column is split by month. Never group a column by itself.

Rules:
- Use column names exactly as listed.
- Blank cells are skipped by sum, average, min, max, top and bottom. Where a table has separate Debit and Credit columns, "debits" simply means the Debit column.
- To count the entries of one kind, use count with that column, or count with a not_blank filter.
- Search words from the question go in a contains filter on the text column that holds them, using the shortest distinctive wording.
- A text column may list what its entries begin with. When the question names a kind of entry (interest, cheque, ATM, charges, salary), filter on the wording the column actually uses for it, taken from that list, not on the question's own word. "interest" may be written "INT.PD", a cheque "CHQ PAID". Use starts_with for a wording from that list.
- If the question names a kind of entry and nothing in the column corresponds to it, return "unsupported". Do not search for a word you have no sign of.
- "What period does this cover" is min and max of the date column.
- Each query is reported on its own. Nothing adds, subtracts or compares two results afterwards. So a question whose answer is arithmetic between values (an opening balance worked back from the first row, a difference between two months, a percentage of a total) is "unsupported". Never return the ingredients in place of the answer.
- Return "unsupported" when the question asks for an explanation, an opinion, a change to the sheet, a formula, one cell by its address, the list of sheets, or anything these operations cannot express. Do not force a plan that only approximates the question.`;

const SAMPLE_VALUES = 3;
const MAX_STARTS = 24;
/** Letters, dots, ampersands and spaces from the start, up to the first digit or separator. */
const LEADING_WORDS = /^[A-Za-z][A-Za-z.&' ]*/;

/**
 * The wordings a text column's entries start with, most common first: "UPI",
 * "ATM WDL", "CHQ PAID", "INT.PD". A statement calls interest "INT.PD" and a
 * cheque "CHQ PAID", so a planner that only knows the question's own words
 * searches for "interest" and finds nothing. These are labels the bank
 * repeats on many rows, not the rows themselves.
 */
export function commonStarts(table: TableModel, column: { index: number }): Array<{ text: string; rows: number }> {
  const counts = new Map<string, { text: string; rows: number }>();
  let filled = 0;
  for (const row of table.rows) {
    const raw = row.cells[column.index];
    if (isBlank(raw)) continue;
    filled += 1;
    const lead = LEADING_WORDS.exec(String(raw).trim())?.[0].trim().split(/\s+/).slice(0, 3).join(' ');
    if (!lead || lead.length < 2) continue;
    const entry = counts.get(lead.toLowerCase()) ?? { text: lead, rows: 0 };
    entry.rows += 1;
    counts.set(lead.toLowerCase(), entry);
  }
  // One wording on every row ("row 2", "row 3") tells the planner nothing, and
  // a column of free text, where almost every entry starts differently, has no
  // vocabulary to show.
  if (counts.size < 2) return [];
  if (filled >= 12 && counts.size > filled * 0.6) return [];
  return [...counts.values()].sort((a, b) => b.rows - a.rows || a.text.localeCompare(b.text)).slice(0, MAX_STARTS);
}

/** The table as the planner sees it: names, types and a few sample cells. Never the full data. */
export function buildTableQueryUserMessage(table: TableModel, message: string): string {
  const lines = table.columns.map((column) => {
    const samples: string[] = [];
    for (const row of table.rows) {
      if (samples.length >= SAMPLE_VALUES) break;
      const raw = row.cells[column.index];
      if (isBlank(raw)) continue;
      const text = column.type === 'date' ? toIsoDateValue(raw) ?? String(raw) : String(raw).trim();
      const clipped = text.length > 48 ? `${text.slice(0, 47)}…` : text;
      if (!samples.includes(clipped)) samples.push(clipped);
    }
    const line = `- "${column.name}" (${column.type})${samples.length ? ` e.g. ${samples.map((s) => JSON.stringify(s)).join(', ')}` : ' (empty)'}`;
    const starts = column.type === 'text' ? commonStarts(table, column) : [];
    return starts.length ? `${line}\n    entries begin with: ${starts.map((s) => `${JSON.stringify(s.text)} x${s.rows}`).join(', ')}` : line;
  });
  return [`Sheet: "${table.sheetName}", ${table.rows.length} data rows.`, 'Columns:', ...lines, '', `Question: ${message.trim()}`].join('\n');
}

export interface TableQuestionInput {
  message: string;
  sheetName: string;
  sheetData: unknown[][];
  headerRowIndex?: number;
  headers?: string[];
  columnHints?: ColumnHint[];
  /** Rows the sheet reports (header included). When more than were read, the answer says so. */
  declaredRowCount?: number;
}

export interface TableAnswer {
  answer: string;
  /** Sheet rows the answer names with "(row N)", for the add-in to point at. */
  rowNumbers: number[];
  /** How many columns the table has, so a pointer can select the whole row. */
  columnCount: number;
}

export class TableQuestionAnswerer {
  private readonly logger = new Logger(TableQuestionAnswerer.name);

  constructor(private readonly openRouter: OpenRouterService) {}

  /**
   * The answer, or null when this question is not one for a query plan (the
   * caller then falls back to its usual path). Never throws.
   */
  async tryAnswer(input: TableQuestionInput): Promise<TableAnswer | null> {
    let table: TableModel;
    try {
      table = buildTable(input);
    } catch (error) {
      this.logger.warn(`Table question: could not read the sheet as a table (${(error as Error).message})`);
      return null;
    }
    if (table.rows.length === 0 || table.columns.length === 0) return null;

    let reply: string;
    const outcome: LlmCompletionOutcome = {};
    try {
      reply = await this.openRouter.complete({
        systemPrompt: TABLE_QUERY_SYSTEM_PROMPT,
        userMessage: buildTableQueryUserMessage(table, input.message),
        tier: 'medium',
        temperature: 0,
        maxTokens: 900,
        reasoningEffort: 'low',
        responseFormat: 'json_object',
        outcome,
      });
    } catch (error) {
      this.logger.warn(`Table question: planner call failed (${(error as Error).message})`);
      return null;
    }
    // A plan cut off by the token cap can still parse while missing queries.
    if (outcome.truncated) {
      this.logger.warn('Table question: planner reply was truncated, discarding it');
      return null;
    }

    try {
      const plan = parseTableQueryPlan(table, reply);
      if (plan.kind === 'unsupported') {
        this.logger.log(`Table question declined by planner: ${plan.reason.slice(0, 120)}`);
        return null;
      }
      const results = plan.queries.map((query) => executeTableQuery(table, query));
      const missing = Math.max(0, (input.declaredRowCount ?? 0) - input.sheetData.length);
      this.logger.log(
        `Table question answered in code: sheet=${table.sheetName} rows=${table.rows.length} queries=${plan.queries.map((q) => q.op).join(',')}`,
      );
      return {
        answer: formatTableAnswer(table, results, { expectedRows: table.rows.length + missing }),
        rowNumbers: rowsMentioned(results),
        columnCount: table.columns.length,
      };
    } catch (error) {
      if (error instanceof TableQueryPlanError) {
        this.logger.warn(`Table question: plan rejected (${error.message})`);
        return null;
      }
      this.logger.error('Table question: execution failed', error as Error);
      return null;
    }
  }
}
