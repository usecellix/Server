import { RouterPath } from '../src/excel-ai/types/router.types';
import { ComplexityTier } from '../src/excel-ai/utils/complexity-classifier.util';

/**
 * The golden set: real-world-shaped prompts mapped to the behavior Cellix is
 * supposed to produce for them. This is the project's reliability baseline —
 * every case here either reproduces a bug that actually shipped (see `regressionOf`)
 * or covers a category of request users send often. The point is to replace
 * "does it feel broken" with a number that moves.
 *
 * Two ways this gets scored:
 * 1. `test/golden-set-eval.spec.ts` — deterministic, free, runs in CI on every
 *    change. Checks the ROUTING/DETECTION layers only (regex fast lanes, the
 *    complexity classifier, header detection) — nothing here calls a real LLM,
 *    so it can't catch bad model output, only bad routing.
 * 2. `eval/run-live-eval.ts` — opt-in, costs real API credits, exercises the
 *    full pipeline including the LLM. Run manually or on a schedule, not on
 *    every commit. See that file for how to run it.
 */

export interface RoutingGoldenCase {
  kind: 'routing';
  id: string;
  category:
    | 'tier0-shortcut'
    | 'tier1-single-action'
    | 'tier2-formula-chart'
    | 'tier3-compound'
    | 'data-query'
    | 'export'
    | 'ask';
  prompt: string;
  mode?: 'ask' | 'action' | 'plan';
  sheetHeaders?: string[];
  expected: {
    route: RouterPath;
    /** Only meaningful for route: 'write'. */
    tier?: ComplexityTier;
    actionHint?: string;
  };
  /** Which specs/ bug report this guards against re-breaking, if any. */
  regressionOf?: string;
}

export interface HeaderDetectionGoldenCase {
  kind: 'header-detection';
  id: string;
  category: 'header-detection';
  description: string;
  sheetData: unknown[][];
  expectedHeaderRowIndex: number;
  expectedHeaders: string[];
  regressionOf?: string;
}

export interface LiveGoldenCase {
  id: string;
  category:
    | 'tier2-formula-chart'
    | 'tier3-compound'
    | 'data-query'
    | 'header-offset-write';
  prompt: string;
  /** Minimal sheet snapshot sent as workbook context. */
  sheetHeaders: string[];
  sheetRows: unknown[][];
  /** Action types that MUST appear in the final action list. */
  mustIncludeActionTypes: string[];
  /** Action types that must NOT appear (e.g. proof a bug doesn't regress). */
  mustNotIncludeActionTypes?: string[];
  regressionOf?: string;
}

export const ROUTING_GOLDEN_SET: RoutingGoldenCase[] = [
  // ---- Tier 0: deterministic, zero LLM ----
  {
    kind: 'routing',
    id: 'tier0-freeze-top-row',
    category: 'tier0-shortcut',
    prompt: 'freeze top row',
    expected: { route: 'shortcut' },
  },
  {
    kind: 'routing',
    id: 'tier0-protect-sheet',
    category: 'tier0-shortcut',
    prompt: 'protect this sheet',
    expected: { route: 'shortcut' },
  },
  {
    kind: 'routing',
    id: 'tier0-bold-range',
    category: 'tier0-shortcut',
    prompt: 'bold cells A1 to C1',
    mode: 'action',
    expected: { route: 'write', tier: 0, actionHint: 'CELL_FORMAT' },
  },
  {
    kind: 'routing',
    id: 'tier0-hide-column',
    category: 'tier0-shortcut',
    prompt: 'hide column F',
    mode: 'action',
    expected: { route: 'write', tier: 0, actionHint: 'VISIBILITY_TOGGLE' },
  },

  // ---- Tier 1: single LLM call ----
  {
    kind: 'routing',
    id: 'tier1-sort-column',
    category: 'tier1-single-action',
    prompt: 'sort column B descending by value',
    mode: 'action',
    expected: { route: 'write', tier: 1, actionHint: 'SORT_OR_FILTER' },
  },
  {
    kind: 'routing',
    id: 'tier1-header-highlight',
    category: 'tier1-single-action',
    prompt: 'highlight the header row in yellow',
    mode: 'action',
    expected: { route: 'write', tier: 1, actionHint: 'HEADER_FORMAT' },
    regressionOf: 'specs/24_header_row_format_matching_rows_leak.md',
  },

  // ---- Tier 2: formulas / charts / pivots — verification mandatory ----
  {
    kind: 'routing',
    id: 'tier2-formula-calc',
    category: 'tier2-formula-chart',
    prompt: 'calculate 18% GST on the amount column',
    mode: 'action',
    expected: { route: 'write', tier: 2, actionHint: 'FORMULA_GEN' },
  },
  {
    kind: 'routing',
    id: 'tier2-chart-request',
    category: 'tier2-formula-chart',
    prompt: 'create a chart of sales by month',
    mode: 'action',
    expected: { route: 'write', tier: 2, actionHint: 'CHART' },
  },

  // ---- Tier 3: compound multi-feature requests ----
  {
    kind: 'routing',
    id: 'tier3-purchase-register-build',
    category: 'tier3-compound',
    prompt:
      'Create a purchase register from the data in this workbook. Add columns for purchase date, ' +
      'supplier, invoice number, item, category, quantity, unit price, tax %, tax amount, total amount, ' +
      'payment status, department, requested by, and approved by. Add formulas for tax and total amount. ' +
      'Add filters, freeze the header row, and create a summary showing total purchases, paid amount, ' +
      'pending amount, and purchases by department.',
    mode: 'action',
    expected: { route: 'write', tier: 3 },
    regressionOf: 'purchase-register-garbage-output (fixed this session — deterministic shortcut misfire)',
  },
  {
    kind: 'routing',
    id: 'tier3-sort-then-chart',
    category: 'tier3-compound',
    prompt: 'sort by column B and then create a chart',
    mode: 'action',
    expected: { route: 'write', tier: 3 },
  },
  {
    kind: 'routing',
    id: 'tier3-dashboard-request',
    category: 'tier3-compound',
    prompt: 'build a dashboard summarizing sales by region with a chart',
    mode: 'action',
    expected: { route: 'write', tier: 3, actionHint: 'DASHBOARD' },
  },
  {
    kind: 'routing',
    id: 'tier3-create-sheet-copy-paid',
    category: 'tier3-compound',
    prompt:
      'create a new sheet named Paid payments and copy the paid data from purchase register to that new sheet only paid',
    mode: 'action',
    expected: { route: 'write' },
    regressionOf: 'specs/20_export_route_misclassification.md',
  },

  // ---- Data query fast lane (read-only) ----
  {
    kind: 'routing',
    id: 'data-query-total',
    category: 'data-query',
    prompt: 'what is the total invoice amount',
    mode: 'ask',
    sheetHeaders: ['Invoice No', 'Supplier', 'Amount'],
    expected: { route: 'data' },
  },
  {
    kind: 'routing',
    id: 'data-query-duplicate-check',
    category: 'data-query',
    prompt: 'are there any duplicate invoice numbers',
    mode: 'ask',
    sheetHeaders: ['Invoice No', 'Supplier', 'Amount'],
    expected: { route: 'data' },
  },

  // ---- Export (find + copy rows to a new sheet) ----
  {
    kind: 'routing',
    id: 'export-find-export-wording',
    category: 'export',
    prompt: 'find Applied and export those rows to a new sheet',
    mode: 'action',
    expected: { route: 'export' },
  },
  {
    kind: 'routing',
    id: 'export-find-copy-wording-overridden-to-write',
    category: 'export',
    prompt: 'find Applied and copy those rows to a new sheet',
    mode: 'action',
    // "copy" is itself a write-intent verb, so the write-intent guard overrides
    // the export-lane match to 'write' — the full Orchestrator pipeline handles
    // find+copy via COPY_FILTERED_RANGE. Correct, existing behavior (see
    // llm-router-instant-shortcut.spec.ts), included here so the two near-
    // identical prompts routing differently stays a documented decision, not
    // a mystery the next person has to re-derive from the regex.
    expected: { route: 'write' },
  },
  {
    kind: 'routing',
    id: 'export-not-misrouted-to-data',
    category: 'export',
    prompt: 'copy all paid invoices to a new sheet',
    mode: 'action',
    expected: { route: 'write' },
    regressionOf: 'llm-router write-intent override (write-intent-guard.util.ts)',
  },

  // ---- Ask mode / read-only explanation ----
  {
    kind: 'routing',
    id: 'ask-mode-readonly',
    category: 'ask',
    prompt: 'what does this workbook contain',
    mode: 'ask',
    expected: { route: 'ask' },
  },
];

export const LIVE_GOLDEN_SET: LiveGoldenCase[] = [
  {
    id: 'live-purchase-register-build',
    category: 'tier3-compound',
    prompt:
      'Create a purchase register from the data in this workbook. Add columns for purchase date, ' +
      'supplier, invoice number, item, category, quantity, unit price, tax %, tax amount, total amount, ' +
      'payment status, department, requested by, and approved by. Add formulas for tax and total amount. ' +
      'Add filters, freeze the header row, and create a summary showing total purchases, paid amount, ' +
      'pending amount, and purchases by department.',
    sheetHeaders: ['row', 'paid amount', 'pending amount', 'purchases by department'],
    sheetRows: [
      ['Value 1', 12500, 12500, 'Value 1'],
      ['Value 2', 28500, 28500, 'Value 2'],
      ['Value 3', 44500, 44500, 'Value 3'],
    ],
    mustIncludeActionTypes: ['INSERT_COLUMN', 'AUTO_FILTER', 'FREEZE_PANES'],
    // The bug this guards against: the deterministic table-shortcut misfiring
    // and returning only WRITE_TABLE with 4 bogus columns and "Value N" filler.
    mustNotIncludeActionTypes: ['WRITE_TABLE'],
    regressionOf: 'purchase-register-garbage-output (fixed this session)',
  },
  {
    id: 'live-sum-formula',
    category: 'tier2-formula-chart',
    prompt: 'add a Total column that is Quantity times Unit Price',
    sheetHeaders: ['Item', 'Quantity', 'Unit Price'],
    sheetRows: [
      ['Widget', 4, 250],
      ['Gadget', 2, 500],
    ],
    mustIncludeActionTypes: ['INSERT_COLUMN'],
  },
  {
    id: 'live-data-query-sum',
    category: 'data-query',
    prompt: 'what is the total amount',
    sheetHeaders: ['Invoice No', 'Supplier', 'Amount'],
    sheetRows: [
      ['INV-001', 'Acme Ltd', 12500],
      ['INV-002', 'Beta Inc', 8500],
    ],
    // Read-only — no actions expected, this checks the answer text separately
    // in a future iteration. For now it's a placeholder proving the route
    // doesn't accidentally mutate the sheet.
    mustIncludeActionTypes: [],
    mustNotIncludeActionTypes: ['SET_CELL', 'SET_FORMULA', 'ADD_ROW'],
  },
  {
    id: 'live-header-offset-freeze',
    category: 'header-offset-write',
    prompt: 'freeze the header row',
    sheetHeaders: ['Date', 'Supplier', 'Amount'],
    // Title row above the real headers — the exact scenario fixed this session.
    // freezeRows should end up targeting the REAL header row, not row 1.
    sheetRows: [
      ['ABC Corp — Purchase Register FY24', '', ''],
      ['', '', ''],
      ['Date', 'Supplier', 'Amount'],
      ['01-04-2024', 'Acme Ltd', 12500],
    ],
    mustIncludeActionTypes: ['FREEZE_PANES'],
    regressionOf: 'header-row-detection width-blindness (fixed this session)',
  },
];

export const HEADER_DETECTION_GOLDEN_SET: HeaderDetectionGoldenCase[] = [
  {
    kind: 'header-detection',
    id: 'header-row1-plain-table',
    category: 'header-detection',
    description: 'Plain table, no preamble — headers on row 1',
    sheetData: [
      ['Name', 'Age', 'City'],
      ['Alice', 30, 'Mumbai'],
    ],
    expectedHeaderRowIndex: 0,
    expectedHeaders: ['Name', 'Age', 'City'],
  },
  {
    kind: 'header-detection',
    id: 'header-single-title-cell-above',
    category: 'header-detection',
    description: 'A single merged title cell above the real headers',
    sheetData: [
      ['ABC Corp — Purchase Register FY24'],
      [],
      ['Date', 'Supplier', 'Invoice No', 'Amount'],
      ['01-04-2024', 'Acme Ltd', 'INV-001', 12500],
    ],
    expectedHeaderRowIndex: 2,
    expectedHeaders: ['Date', 'Supplier', 'Invoice No', 'Amount'],
    regressionOf: 'header-row-detection width-blindness (fixed this session)',
  },
  {
    kind: 'header-detection',
    id: 'header-title-and-subtitle-above',
    category: 'header-detection',
    description: 'Title row + subtitle row + blank row above headers (Tally-style export)',
    sheetData: [
      ['Purchase Register'],
      ['For the year ended 31 March 2024'],
      [],
      ['Date', 'Supplier', 'Invoice No', 'Amount', 'Tax %', 'Total'],
      ['01-04-2024', 'Acme Ltd', 'INV-001', 12500, 18, 14750],
    ],
    expectedHeaderRowIndex: 3,
    expectedHeaders: ['Date', 'Supplier', 'Invoice No', 'Amount', 'Tax %', 'Total'],
    regressionOf: 'header-row-detection width-blindness (fixed this session)',
  },
];
