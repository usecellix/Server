import { CellChange } from '../src/audit/types/change-set.types';
import { SheetAction } from '../src/excel-ai/types/sheet-actions.types';
import {
  buildInternalDetails,
  buildUserFacingSummary,
  describeRangeCompactly,
  INTERNAL_COPY_MARKERS,
  sanitizeAnswerForHeadline,
  tierProcessingLabel,
  summarizePlanIntent,
  type PlanIntent,
} from '../src/excel-ai/utils/user-facing-response.util';

function cell(sheet: string, address: string): CellChange {
  return {
    cell: address,
    sheet,
    before: null,
    after: 1,
    isHardcoded: true,
  };
}

describe('user-facing-response.util', () => {
  describe('describeRangeCompactly', () => {
    it('collapses a rectangular block to a compact A1 range', () => {
      const changes = [
        cell('Purchase Register', 'A53'),
        cell('Purchase Register', 'B53'),
        cell('Purchase Register', 'C53'),
        cell('Purchase Register', 'A54'),
        cell('Purchase Register', 'B54'),
        cell('Purchase Register', 'C54'),
        cell('Purchase Register', 'A55'),
        cell('Purchase Register', 'B55'),
        cell('Purchase Register', 'C55'),
      ];
      expect(describeRangeCompactly(changes)).toBe('Purchase Register!A53:C55');
    });
  });

  describe('buildUserFacingSummary — 9-cell / 1-action regression', () => {
    it('derives supporting detail from cell diffs, not action count', () => {
      const actions: SheetAction[] = [
        {
          type: 'AGGREGATE_TABLE',
          sheetName: 'Purchase Register',
          groupByColumn: 'Payment Status',
        },
      ];
      const changes = [
        cell('Purchase Register', 'A53'),
        cell('Purchase Register', 'B53'),
        cell('Purchase Register', 'C53'),
        cell('Purchase Register', 'A54'),
        cell('Purchase Register', 'B54'),
        cell('Purchase Register', 'C54'),
        cell('Purchase Register', 'A55'),
        cell('Purchase Register', 'B55'),
        cell('Purchase Register', 'C55'),
      ];

      const summary = buildUserFacingSummary({
        answer: 'Added a Payment Status summary.',
        actions,
        changes,
        activeSheetName: 'Purchase Register',
      });

      expect(summary.contextLine).toBe('Working with: Purchase Register');
      expect(summary.headline).toBe('Added a Payment Status summary.');
      expect(summary.supportingDetail).toBe('9 cells, Purchase Register!A53:C55');
      expect(summary.bullets).toBeUndefined();
      expect(INTERNAL_COPY_MARKERS.test(summary.headline)).toBe(false);
    });
  });

  describe('assumption / typo fixture', () => {
    it('states the assumption in the visible headline', () => {
      const summary = buildUserFacingSummary({
        answer: "I'll sort with Paid first, then Pending.",
        actions: [{ type: 'SORT_RANGE', sheetName: 'Purchase Register', key: 10, ascending: true }],
        assumption:
          "I noticed 'paid should be first then paid' looks like it may have a typo — I've sorted with Paid first, then Pending",
        activeSheetName: 'Purchase Register',
      });

      expect(summary.headline.toLowerCase()).toContain('typo');
      expect(summary.headline.toLowerCase()).toContain('paid first');
      expect(INTERNAL_COPY_MARKERS.test(summary.headline)).toBe(false);
    });
  });

  describe('internal copy never in headline', () => {
    it('strips tier jargon from answer and falls back to action description', () => {
      expect(
        sanitizeAnswerForHeadline(
          'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
        ),
      ).toBe('');

      const summary = buildUserFacingSummary({
        answer: 'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
        actions: [
          {
            type: 'FORMAT_MATCHING_ROWS',
            sheetName: 'Purchase Register',
            format: { fillColor: '#FFEB9C' },
          },
        ],
      });

      expect(INTERNAL_COPY_MARKERS.test(summary.headline)).toBe(false);
      expect(summary.headline.toLowerCase()).toContain('highlight');
    });
  });

  describe('buildInternalDetails', () => {
    it('preserves legacy tier explanation and raw actions', () => {
      const details = buildInternalDetails({
        tier: 1,
        model: 'openai/gpt-5-mini',
        processingLabel: tierProcessingLabel(1, 'CONDITIONAL_FORMAT'),
        actions: [{ type: 'FORMAT_MATCHING_ROWS', range: 'A2:L51', sheetName: 'Purchase Register' }],
        legacyExplanation: tierProcessingLabel(1, 'CONDITIONAL_FORMAT'),
        assumption: 'Paid first, then Pending',
      });

      expect(details.processingLabel).toContain('Tier 1');
      expect(details.model).toBe('openai/gpt-5-mini');
      expect(details.rawActionSummary).toContain('FORMAT_MATCHING_ROWS');
      expect(details.assumption).toContain('Paid first');
    });
  });
});

/**
 * TASKS.md #140 — the action card grew into a wall of text.
 *
 * A 13-sheet build produced ~40 distinct action descriptions ("Freeze Panes on
 * January", "Freeze Panes on February", …) and `buildUserFacingSummary` put
 * every one of them in `bullets`, which the card renders in its body. The card
 * became taller than the panel and pushed its own Accept button off-screen.
 */
describe('buildUserFacingSummary — large multi-sheet batches stay summarizable', () => {
  const MONTHS = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
  ];

  function bigBatch(): SheetAction[] {
    const actions: SheetAction[] = [];
    for (const sheet of [...MONTHS, 'Main']) {
      actions.push({ type: 'SET_CELL', sheetName: sheet, row: 0, col: 0, value: 'Unit No' });
      actions.push({ type: 'FREEZE_PANES', sheetName: sheet, freezeRows: 1 });
      actions.push({ type: 'AUTOFIT_COLUMNS', sheetName: sheet });
    }
    return actions;
  }

  it('rolls per-sheet repetition up into one line per kind of change', () => {
    const summary = buildUserFacingSummary({ actions: bigBatch(), changes: [] });
    expect(summary.bullets).toBeDefined();
    expect(summary.bullets!.length).toBeLessThanOrEqual(6);
    expect(summary.bullets).toContain('Update cell values on 13 sheets');
    expect(summary.bullets).toContain('Freeze Panes on 13 sheets');
    expect(summary.bullets).toContain('Autofit Columns on 13 sheets');
  });

  it('never lists one bullet per sheet', () => {
    const summary = buildUserFacingSummary({ actions: bigBatch(), changes: [] });
    expect(summary.bullets!.some((b) => b.includes('January'))).toBe(false);
  });

  it('leaves small batches enumerated as before', () => {
    const summary = buildUserFacingSummary({
      actions: [
        { type: 'SET_CELL', sheetName: 'Main', row: 1, col: 0, value: 'A' },
        { type: 'SORT_RANGE', sheetName: 'Main', range: 'A1:C9' },
      ],
      changes: [],
    });
    expect(summary.bullets).toEqual(['Update cell values on Main', 'Sort the sheet on Main']);
  });

  it('drops bullets entirely when even the rolled-up list is too long', () => {
    const actions: SheetAction[] = [
      { type: 'SET_CELL', sheetName: 'S', row: 1, col: 0, value: 'x' },
      { type: 'SORT_RANGE', sheetName: 'S', range: 'A1:B2' },
      { type: 'CREATE_CHART', sheetName: 'S', sourceRange: 'A1:B2' },
      { type: 'DELETE_ROW', sheetName: 'S', row: 2 },
      { type: 'ADD_ROW', sheetName: 'S', data: ['a'] },
      { type: 'FREEZE_PANES', sheetName: 'S', freezeRows: 1 },
      { type: 'AUTOFIT_COLUMNS', sheetName: 'S' },
      { type: 'PROTECT_SHEET', sheetName: 'S' },
    ];
    const summary = buildUserFacingSummary({ actions, changes: [] });
    // 8 distinct kinds on one sheet — nothing to roll up, so the card shows the
    // headline only and the details disclosure carries the list.
    expect(summary.bullets).toBeUndefined();
  });
});

/**
 * TASKS.md #149 — the Accept card describes intent, not mechanics.
 *
 * Fixtures are the real subtask descriptions from a live monthly-ledger run
 * (19 subtasks), not invented shapes — the lesson from #143/#144 was that
 * fixtures must come from what the planner actually emits.
 */
describe('summarizePlanIntent — plan intent replaces action mechanics', () => {
  const MONTHS = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
  ];

  const HEADER_COLS =
    '[Unit No, Guest, Guest Name, Check In, Check Out, Rate Per Night, Total Amount, Source, Payment Status, Bank Account]';

  function realLedgerPlan(): PlanIntent[] {
    const subtasks: PlanIntent[] = [
      { id: 's1', description: "Create sheet 'Main' if it doesn't exist", targetSheet: 'Main' },
    ];
    MONTHS.forEach((m, i) =>
      subtasks.push({
        id: `m${i + 1}`,
        description: `Create sheet '${m}' if it doesn't exist and set A1:J1 headers ${HEADER_COLS}`,
        targetSheet: m,
      }),
    );
    subtasks.push({
      id: 's2',
      description:
        'On Main, set A4:D4 headers to [Month, Total Amount, Paid Amount, Pending Amount] and fill A5:A16 with the month names',
      targetSheet: 'Main',
    });
    subtasks.push({
      id: 's3',
      description:
        'Write Consolidated Transactions header at Main!A18 with columns [Month, Unit No, Guest, Guest Name, Check In, Check Out, Rate Per Night, Total Amount, Source, Payment Status, Bank Account]',
      targetSheet: 'Main',
    });
    return subtasks;
  }

  it('collapses twelve identically-shaped month subtasks into one counted line', () => {
    const bullets = summarizePlanIntent(realLedgerPlan());
    const monthLine = bullets.find((b) => b.includes('12 sheets'));
    expect(monthLine).toBeDefined();
    expect(monthLine).toContain('January, February, March +9 more');
  });

  it('takes a 15-subtask plan down to a readable handful', () => {
    const bullets = summarizePlanIntent(realLedgerPlan());
    // 15 subtasks in, 4 distinct intents out (Main create, 12 months, totals
    // headers, consolidated header).
    expect(bullets.length).toBeLessThanOrEqual(6);
    expect(bullets.length).toBeGreaterThanOrEqual(3);
  });

  it('never names an individual month as if the line applied only to it', () => {
    const bullets = summarizePlanIntent(realLedgerPlan());
    const monthLine = bullets.find((b) => b.includes('12 sheets'))!;
    // "Create sheet 'January' ..." must not survive as the group's label.
    expect(monthLine).not.toMatch(/Create sheet 'January'/);
  });

  it('keeps genuinely distinct subtasks separate', () => {
    const bullets = summarizePlanIntent(realLedgerPlan());
    expect(bullets.some((b) => b.includes('Consolidated Transactions'))).toBe(true);
    expect(bullets.some((b) => b.includes('Total Amount, Paid Amount'))).toBe(true);
  });

  it('trims a subtask whose description is a wall of formulas', () => {
    const long = Array.from({ length: 18 }, (_, i) => `B${i + 5} =SUM(January!G:G)`).join(', ');
    const bullets = summarizePlanIntent([
      { id: 's1', description: `Write Monthly Totals formulas: ${long}`, targetSheet: 'Main' },
      { id: 's2', description: 'Create a chart', targetSheet: 'Main' },
    ]);
    expect(bullets[0].length).toBeLessThanOrEqual(160);
    expect(bullets[0].endsWith('...')).toBe(true);
  });

  it('is generic — groups by targetSheet, not by month names', () => {
    // The same collapse must work for any repeated shape, which is what long
    // prompts produce in general.
    const regions = ['North', 'South', 'East', 'West'];
    const bullets = summarizePlanIntent(
      regions.map((r, i) => ({
        id: `r${i}`,
        description: `Create sheet '${r}' and set A1:C1 headers [Rep, Target, Actual]`,
        targetSheet: r,
      })),
    );
    expect(bullets).toHaveLength(1);
    expect(bullets[0]).toContain('4 sheets');
    expect(bullets[0]).toContain('North, South, East +1 more');
  });

  it('handles sheet names containing regex metacharacters', () => {
    const bullets = summarizePlanIntent([
      { id: 'a', description: "Create sheet 'Q1 (2026)' and add headers", targetSheet: 'Q1 (2026)' },
      { id: 'b', description: "Create sheet 'Q2 (2026)' and add headers", targetSheet: 'Q2 (2026)' },
    ]);
    expect(bullets).toHaveLength(1);
    expect(bullets[0]).toContain('2 sheets');
  });

  it('returns nothing for an empty or missing plan', () => {
    expect(summarizePlanIntent([])).toEqual([]);
    expect(summarizePlanIntent(undefined as unknown as PlanIntent[])).toEqual([]);
  });

  it('buildUserFacingSummary prefers plan intent over the action rollup', () => {
    const actions: SheetAction[] = MONTHS.flatMap((m) => [
      { type: 'SET_CELL', sheetName: m, row: 0, col: 0, value: 'Unit No' },
      { type: 'FREEZE_PANES', sheetName: m, freezeRows: 1 },
      { type: 'AUTOFIT_COLUMNS', sheetName: m },
    ]);

    const withPlan = buildUserFacingSummary({
      actions,
      changes: [],
      planSubtasks: realLedgerPlan(),
    });
    const withoutPlan = buildUserFacingSummary({ actions, changes: [] });

    // With a plan: intent. Without: the mechanical rollup (Tier 0-2 path).
    expect(withPlan.bullets!.some((b) => b.includes('12 sheets'))).toBe(true);
    expect(withoutPlan.bullets!.some((b) => b.startsWith('Freeze Panes'))).toBe(true);
  });
});
