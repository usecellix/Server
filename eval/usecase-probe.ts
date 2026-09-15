/**
 * Live probe for every use case in `cellix-basic-usecases.html`.
 *
 * Sibling of `run-live-eval.ts`: same "send real prompts at a running backend"
 * shape, but the case list is the guide's own phrasing, section by section, and
 * the checks encode what the guide promises (which action types must appear,
 * which must NOT, and whether the request is read-only). It costs real
 * OpenRouter credits and is not part of `npm test`.
 *
 * Usage:
 *   1. CELLIX_EVAL_BYPASS_TOKEN=local-eval npm run start:dev
 *   2. CELLIX_EVAL_BYPASS_TOKEN=local-eval npx ts-node eval/usecase-probe.ts
 *      Filter to specific cases:  ONLY=sheet-copy,filter-gt
 *      Filter to a section:       SECTION=T1.1
 *      Point elsewhere:           CELLIX_EVAL_BASE_URL=http://localhost:4011
 *
 * Results land in eval/usecase-results.json (git-ignored scratch) with the full
 * answer text and actions for each case, so a regression can be diffed rather
 * than re-read.
 */
import * as fs from 'fs';
import * as path from 'path';
import { buildFixture } from './usecase-fixture';

const BASE_URL = process.env.CELLIX_EVAL_BASE_URL ?? 'http://localhost:4001';
const TOKEN = process.env.CELLIX_EVAL_BYPASS_TOKEN ?? 'usecase-probe';
const ONLY = (process.env.ONLY ?? '').split(',').map((s) => s.trim()).filter(Boolean);
const SECTION = process.env.SECTION ?? '';
const OUT = process.env.OUT ?? path.join(__dirname, 'usecase-results.json');
const CASE_TIMEOUT_MS = Number(process.env.CASE_TIMEOUT_MS ?? 240_000);

export interface UseCase {
  id: string;
  /** Section in cellix-basic-usecases.html this comes from. */
  section: string;
  prompt: string;
  /** Action types that must appear; "A|B" means either satisfies it. "*" means "any action". */
  expect?: string[];
  /** Action types that must NOT appear — the guide's own safety expectations. */
  forbid?: string[];
  /** Guide marks it read-only (Q&A): an answer with zero actions is the pass. */
  readOnly?: boolean;
}

const fixture = buildFixture();
const LAST = fixture.lastRow;

export const USE_CASES: UseCase[] = [
  // ── T1.1 Sheet operations ─────────────────────────────────────────────
  { id: 'sheet-create', section: 'T1.1', prompt: 'Create a new sheet called Q1 Summary', expect: ['ADD_SHEET|CREATE_SHEET'] },
  { id: 'sheet-rename', section: 'T1.1', prompt: 'Rename this sheet to Apr 2024 Data', expect: ['RENAME_SHEET'] },
  { id: 'sheet-copy', section: 'T1.1', prompt: 'Copy the Purchase Register sheet and name it March Copy', expect: ['COPY_SHEET|ADD_SHEET'] },
  { id: 'sheet-delete', section: 'T1.1', prompt: 'Delete the Summary sheet', expect: ['DELETE_SHEET'] },
  { id: 'sheet-hide-named', section: 'T1.1', prompt: 'Hide the GSTR-2A sheet', expect: ['HIDE_SHEET'] },
  { id: 'sheet-unhide', section: 'T1.1', prompt: 'Unhide the Working sheet', expect: ['SHOW_SHEET'], forbid: ['HIDE_SHEET'] },
  { id: 'sheet-tabcolor', section: 'T1.1', prompt: 'Change the Summary tab colour to blue', expect: ['SET_SHEET_COLOR'] },
  { id: 'sheet-move', section: 'T1.1', prompt: 'Move the Summary sheet to the first position', expect: ['*'], forbid: ['DELETE_SHEET'] },
  { id: 'sheet-count', section: 'T1.1', prompt: 'How many sheets are in this workbook?', readOnly: true },

  // ── T1.2 Row & column operations ──────────────────────────────────────
  { id: 'row-insert', section: 'T1.2', prompt: 'Insert 3 blank rows after row 5', expect: ['INSERT_ROW|ADD_ROW'] },
  { id: 'row-delete-blank-gstin', section: 'T1.2', prompt: 'Delete all rows where column D is blank', expect: ['DELETE_MATCHING_ROWS|DELETE_ROW'], forbid: ['DELETE_SHEET'] },
  // The fixture has NO fully blank rows, so the only correct outcome is to say
  // so and emit nothing. Live runs produced DELETE_ROW row:10 rowCount:21 and
  // DELETE_ROW row:1 rowCount:30 — i.e. offers to delete 21 and then all 30
  // rows of real data, both marked "verified: true". TASKS.md #234.
  { id: 'row-delete-blank-sheet', section: 'T1.2', prompt: 'Delete blank rows in the Purchase Register sheet', expect: ['DELETE_MATCHING_ROWS'], forbid: ['DELETE_SHEET', 'DELETE_ROW'] },
  { id: 'col-delete-in-sheet', section: 'T1.2', prompt: 'Delete column C from this sheet', expect: ['DELETE_COLUMN'], forbid: ['DELETE_SHEET'] },
  { id: 'freeze-top', section: 'T1.2', prompt: 'Freeze the top row', expect: ['FREEZE_PANES'] },
  { id: 'autofit', section: 'T1.2', prompt: 'Auto-fit all column widths', expect: ['AUTOFIT_COLUMNS'] },
  { id: 'hide-col-name', section: 'T1.2', prompt: 'Hide the Narration column', expect: ['HIDE_COLUMN'] },
  { id: 'row-height-range', section: 'T1.2', prompt: 'Set the height of rows 2 to 5 to 25', expect: ['SET_ROW_HEIGHT'] },
  { id: 'row-count', section: 'T1.2', prompt: 'How many rows of data do I have?', readOnly: true },

  // ── T1.3 Cell operations ──────────────────────────────────────────────
  { id: 'cell-read', section: 'T1.3', prompt: 'What is in cell E3?', readOnly: true },
  { id: 'cell-merge', section: 'T1.3', prompt: 'Merge and centre A1:C1', expect: ['MERGE_CELLS'] },
  { id: 'cell-clear-format', section: 'T1.3', prompt: `Clear all formatting in the range A1:I${LAST}`, expect: ['*'] },
  { id: 'cell-clear-column', section: 'T1.3', prompt: 'Clear all data in column C', expect: ['*'], forbid: ['CLEAR_ALL'] },
  { id: 'cell-comment', section: 'T1.3', prompt: 'Add a comment to E9 saying "Check this amount"', expect: ['ADD_COMMENT'] },

  // ── T1.4 Formatting ───────────────────────────────────────────────────
  { id: 'fmt-inr', section: 'T1.4', prompt: 'Format column E as Indian rupee currency', expect: ['FORMAT_RANGE'] },
  { id: 'fmt-header', section: 'T1.4', prompt: 'Make the header row bold with a grey background', expect: ['FORMAT_RANGE'] },
  { id: 'fmt-date', section: 'T1.4', prompt: 'Format all dates in column B as dd-mm-yyyy', expect: ['FORMAT_RANGE'] },
  { id: 'fmt-border', section: 'T1.4', prompt: `Add a thick box border around A1:I${LAST}`, expect: ['FORMAT_RANGE'] },

  // ── T1.5 Copy / paste / fill ──────────────────────────────────────────
  { id: 'fill-serial', section: 'T1.5', prompt: 'Insert a column at the start and fill serial numbers 1 to 30 in it', expect: ['*'], forbid: ['WRITE_TABLE'] },
  { id: 'fill-values', section: 'T1.5', prompt: 'Replace all formulas in column F with their values', expect: ['*'] },

  // ── T1.6 Math formulas ────────────────────────────────────────────────
  { id: 'math-total', section: 'T1.6', prompt: `Calculate the total IGST in column F and put it in F${LAST + 2}`, expect: ['SET_FORMULA|SET_CELL|BATCH_SET'] },
  { id: 'math-max', section: 'T1.6', prompt: 'What is the highest taxable amount in this sheet?', readOnly: true },
  { id: 'math-gst', section: 'T1.6', prompt: 'Add a column that calculates the GST inclusive amount as taxable amount times 1.18', expect: ['*'], forbid: ['WRITE_TABLE'] },

  // ── T1.7 Logical formulas ─────────────────────────────────────────────
  { id: 'logic-missing-gstin', section: 'T1.7', prompt: "If the GSTIN in column D is blank, mark a new Status column as 'Missing GSTIN'", expect: ['*'] },
  { id: 'logic-high-value', section: 'T1.7', prompt: "Add a column: if taxable amount is above 1 lakh mark 'High Value', else 'Standard'", expect: ['*'] },

  // ── T2.1 / T2.2 / T2.3 sort, filter, find ─────────────────────────────
  { id: 'sort-date', section: 'T2.1', prompt: 'Sort this table by invoice date, newest first', expect: ['SORT_RANGE'] },
  { id: 'sort-multi', section: 'T2.1', prompt: 'Sort by supplier name A-Z, then by taxable amount highest to lowest', expect: ['SORT_RANGE'] },
  { id: 'filter-gt', section: 'T2.2', prompt: 'Show only rows where the taxable amount is above 1 lakh', expect: ['*'], forbid: ['SORT_RANGE'] },
  { id: 'filter-clear', section: 'T2.2', prompt: 'Clear all filters', expect: ['*'] },
  { id: 'filter-state', section: 'T2.2', prompt: 'What filter is active right now?', readOnly: true },
  { id: 'replace-text', section: 'T2.3', prompt: "Replace all occurrences of 'ABC Traders' with 'ABC Traders Pvt Ltd'", expect: ['*'] },
  { id: 'find-prefix', section: 'T2.3', prompt: "Find all rows with the supplier name starting with 'Kerala'", readOnly: true },

  // ── T2.4–T2.7 formulas ────────────────────────────────────────────────
  { id: 'lookup-state', section: 'T2.4', prompt: 'Add a State Code column that extracts the first 2 digits of the GSTIN', expect: ['*'] },
  { id: 'text-trim', section: 'T2.5', prompt: 'Trim the trailing spaces from the supplier names in column C', expect: ['*'] },
  { id: 'sumif-kerala', section: 'T2.7', prompt: 'What is the total purchase amount from Kerala suppliers (GSTIN starting 32)?', readOnly: true },
  { id: 'summary-per-supplier', section: 'T2.7', prompt: 'Create a summary in a new sheet showing total taxable amount per supplier', expect: ['AGGREGATE_TABLE|WRITE_TABLE|BATCH_SET|SET_FORMULA'] },

  // ── T2.8 Conditional formatting ───────────────────────────────────────
  { id: 'cf-rows-red', section: 'T2.8', prompt: 'Highlight all rows in red where the taxable amount is above 5 lakhs', expect: ['CONDITIONAL_FORMAT|FORMAT_MATCHING_ROWS'] },
  { id: 'cf-duplicates', section: 'T2.8', prompt: 'Highlight duplicate invoice numbers in column A', expect: ['CONDITIONAL_FORMAT|FORMAT_MATCHING_ROWS|HIGHLIGHT_CELL|FORMAT_RANGE'] },
  { id: 'cf-colorscale', section: 'T2.8', prompt: 'Add a colour scale to the taxable amount column — green for low, red for high', expect: ['CONDITIONAL_FORMAT'] },
  { id: 'cf-databars', section: 'T2.8', prompt: 'Add data bars to the taxable amount column', expect: ['CONDITIONAL_FORMAT'] },
  { id: 'cf-clear', section: 'T2.8', prompt: 'Clear all conditional formatting from this sheet', expect: ['*'] },

  // ── T3 ────────────────────────────────────────────────────────────────
  { id: 'err-scan', section: 'T3.1', prompt: 'Scan this sheet and tell me all the errors', readOnly: true },
  { id: 'val-dropdown', section: 'T3.2', prompt: 'Add a dropdown in a new Tax Type column with options: IGST, CGST+SGST, Exempt', expect: ['DATA_VALIDATION'] },
  { id: 'val-length', section: 'T3.2', prompt: 'Add a validation rule so GSTINs in column D must be exactly 15 characters', expect: ['DATA_VALIDATION'] },
  { id: 'named-create', section: 'T3.3', prompt: `Name the range E2:E${LAST} as TaxableAmount`, expect: ['DEFINE_NAMED_RANGE'] },
  { id: 'named-list', section: 'T3.3', prompt: 'What named ranges exist in this workbook?', readOnly: true },
  { id: 'dedupe-remove', section: 'T3.5', prompt: 'Find and remove duplicate invoice numbers in column A', expect: ['DELETE_MATCHING_ROWS|DELETE_ROW|SET_RANGE_VALUES|BATCH_SET'] },
  { id: 'text-to-number', section: 'T3.5', prompt: 'Column E has amounts stored as text — convert them to numbers', expect: ['*'] },
  { id: 'pivot', section: 'T3.6', prompt: 'Create a pivot table showing total amount by supplier', expect: ['AGGREGATE_TABLE|WRITE_TABLE'] },
  { id: 'chart-bar', section: 'T3.7', prompt: 'Create a bar chart showing total taxable amount by supplier', expect: ['CREATE_CHART'] },

  // ── Q&A layer ─────────────────────────────────────────────────────────
  { id: 'qa-describe', section: 'Q&A.2', prompt: 'Describe this spreadsheet to me', readOnly: true },
  { id: 'qa-explain-formula', section: 'Q&A.2', prompt: 'What does the formula in F2 do?', readOnly: true },
  { id: 'qa-merges', section: 'Q&A.2', prompt: 'Are there any merged cells in this sheet?', readOnly: true },
  { id: 'qa-top-supplier', section: 'Q&A.4', prompt: 'Which supplier has the highest total taxable amount?', readOnly: true },
  { id: 'qa-blank', section: 'Q&A.4', prompt: 'How many blank cells are in the GSTIN column?', readOnly: true },
  { id: 'qa-dupes', section: 'Q&A.4', prompt: 'Are there duplicate values in column A?', readOnly: true },
  { id: 'qa-sum-zero', section: 'Q&A.3', prompt: 'Column E amounts are not summing correctly — why?', readOnly: true },
];

interface CaseResult extends UseCase {
  ms: number;
  http: number;
  types: string[];
  actions: Record<string, unknown>[];
  text: string;
  errors: string[];
  statuses: string[];
  verdict: string;
}

async function runCase(testCase: UseCase): Promise<CaseResult> {
  const startedAt = Date.now();
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), CASE_TIMEOUT_MS);

  const actions: Record<string, unknown>[] = [];
  const errors: string[] = [];
  const statuses: string[] = [];
  let text = '';
  let http = 0;

  try {
    const response = await fetch(`${BASE_URL}/excel-ai/conversation`, {
      method: 'POST',
      signal: controller.signal,
      headers: { 'Content-Type': 'application/json', 'x-cellix-eval-bypass': TOKEN },
      body: JSON.stringify({
        message: testCase.prompt,
        sheetData: fixture.sheetData,
        workbookContext: fixture.workbookContext,
        mode: 'action',
        previewEnabled: true,
        excelCapabilities: { dynamicArrays: true, probed: true },
      }),
    });
    http = response.status;

    if (!response.ok || !response.body) {
      errors.push(`HTTP ${response.status}: ${(await response.text()).slice(0, 200)}`);
    } else {
      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const frames = buffer.split('\n\n');
        buffer = frames.pop() ?? '';
        for (const frame of frames) {
          const event = frame.split('\n').find((l) => l.startsWith('event:'))?.slice(6).trim();
          const data = frame
            .split('\n')
            .filter((l) => l.startsWith('data:'))
            .map((l) => l.slice(5).trim())
            .join('');
          if (!event) continue;
          let parsed: Record<string, any> = {};
          try {
            parsed = JSON.parse(data);
          } catch {
            /* partial frame — the case fails on its checks instead */
          }
          if (event === 'chunk' && typeof parsed.text === 'string') text += parsed.text;
          else if (event === 'actions') {
            actions.push(...((parsed.actions as Record<string, unknown>[]) ?? []));
            if (parsed.explanation) text += `\n[explanation] ${parsed.explanation}`;
          } else if (event === 'error') errors.push(String(parsed.message ?? data));
          else if (event === 'status' && parsed.message) statuses.push(String(parsed.message));
          else if (['answer', 'plan_only', 'clarification'].includes(event)) {
            text += `\n[${event}] ${String(parsed.answer ?? parsed.text ?? parsed.message ?? data)}`;
          }
        }
      }
    }
  } catch (err) {
    errors.push((err as Error)?.name === 'AbortError' ? `TIMEOUT ${CASE_TIMEOUT_MS}ms` : String(err));
  }
  clearTimeout(timer);

  const types = actions.map((a) => String(a.type));
  const missing = (testCase.expect ?? [])
    .filter((want) => want !== '*')
    .filter((want) => !want.split('|').some((t) => types.includes(t)));
  const forbidden = (testCase.forbid ?? []).filter((t) => types.includes(t));

  let verdict = 'PASS';
  if (errors.length) verdict = `ERROR: ${errors[0].slice(0, 80)}`;
  else if (forbidden.length) verdict = `FORBIDDEN ${forbidden.join(',')}`;
  else if (missing.length) verdict = `MISSING ${missing.join(',')}`;
  else if (testCase.expect && types.length === 0) verdict = 'NO ACTIONS';
  else if (testCase.readOnly && types.length > 0) verdict = `WROTE ON READ-ONLY (${types.join(',')})`;

  return {
    ...testCase,
    ms: Date.now() - startedAt,
    http,
    types,
    actions,
    text: text.trim().slice(0, 4000),
    errors,
    statuses: statuses.slice(-4),
    verdict,
  };
}

async function main(): Promise<void> {
  const cases = USE_CASES.filter(
    (c) => (ONLY.length === 0 || ONLY.includes(c.id)) && (!SECTION || c.section === SECTION),
  );
  console.log(`Running ${cases.length} use-case probe(s) against ${BASE_URL}\n`);

  const results: CaseResult[] = [];
  for (const testCase of cases) {
    process.stdout.write(`  ${testCase.id.padEnd(24)} `);
    const result = await runCase(testCase);
    results.push(result);
    fs.writeFileSync(OUT, JSON.stringify(results, null, 2));
    console.log(
      `${result.verdict.padEnd(26)} ${String(result.ms).padStart(6)}ms  [${result.types.join(', ')}]`,
    );
  }

  const passed = results.filter((r) => r.verdict === 'PASS');
  console.log(`\n=== ${passed.length}/${results.length} passed — details in ${OUT} ===`);
  for (const result of results.filter((r) => r.verdict !== 'PASS')) {
    console.log(`  ${result.verdict}  ${result.id} [${result.section}] ${JSON.stringify(result.prompt)}`);
  }
  process.exit(passed.length === results.length ? 0 : 1);
}

if (require.main === module) {
  main().catch((err) => {
    console.error('Use-case probe crashed:', err);
    process.exit(1);
  });
}
