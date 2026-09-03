import { PlannerAgent } from '../src/agents/planner.agent';
import { PlannerExhaustedError } from '../src/agents/errors';
import {
  PLANNER_LAST_RESORT_MAX_TOKENS,
  resolvePlannerMaxTokens,
  resolveTier3ComplexityScore,
} from '../src/agents/utils/planner-token-budget.util';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { PlannerOutput, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Spec 16 fix #3 — the fallback-plan path (a single subtask whose description is
 * the raw user message, or the PlannerExhaustedError it was replaced by) must be
 * unreachable in normal operation once the completion-token ceiling is scaled by
 * complexity score. These fixtures stand in for `cellix-benchmark-test-prompts.md`
 * Tier 3-4 prompts (that file does not exist in this repo checkout) — each is a
 * realistic compound/dashboard-style request of the kind spec 16's repro came from.
 */
const TIER3_4_BENCHMARK_PROMPTS: string[] = [
  // Spec 16's exact repro.
  'In dashboard create a chart ,and analysis for purchase register a summary for purchase register',
  'Build a dashboard with a pivot table, a chart, and a KPI summary',
  'Create a Dashboard sheet with a chart of monthly sales and a summary table',
  'Add a sheet for each month with a ledger table, then a Main sheet with a consolidated summary and a chart',
  'Build a purchase register dashboard: a summary, an analysis section, a pivot table, and a chart of totals',
  'build a fancy dashboard with charts and summaries',
  // Follow-up regression: named no TIER3_OBJECT_KEYWORDS hit at all before the
  // keyword list was widened (trend/total/region/breakdown/comparison/month-
  // quarter-year), scoring 1 and getting only the 4096 base budget despite
  // clearly asking for a multi-cut breakdown.
  'build me something showing trends by region and by month, plus totals',
];

function sheet(
  name: string,
  values: unknown[][],
  rowCount = values.length,
  columnCount = values[0]?.length ?? 0,
): WorkbookContext['sheets'][number] {
  return {
    name,
    usedRange: rowCount && columnCount ? `A1:${String.fromCharCode(64 + columnCount)}${rowCount}` : 'A1',
    rowCount,
    columnCount,
    values,
    formulas: values.map((row) => row.map(() => '')),
    numberFormats: values.map((row) => row.map(() => 'General')),
    structure: 'data_table',
    headerRowIndex: 0,
  };
}

function nonEmptyContext(): WorkbookContext {
  return {
    activeSheetName: 'Dashboard',
    sheets: [
      sheet('Dashboard', [
        ['Col A', 'Col B'],
        [1, 2],
      ]),
      sheet('Purchase Register', [
        ['Item', 'Amount'],
        ['Widget', 100],
      ]),
    ],
    namedRanges: [],
    tables: [],
  };
}

function buildAgent(completeImpl: jest.Mock): PlannerAgent {
  const llm = { complete: completeImpl } as unknown as OpenRouterService;
  const config = {
    openRouterModelHigh: 'openai/gpt-5',
    openRouterModelPlanner: 'openai/gpt-5',
  } as unknown as AppConfigService;
  return new PlannerAgent(llm, config);
}

/** A minimal-but-valid plan whose subtask count scales with the prompt's complexity
 * score, standing in for the real plan a model would emit for that request. */
function planFor(prompt: string): PlannerOutput {
  const score = resolveTier3ComplexityScore(prompt);
  const subtasks = Array.from({ length: score }, (_, i) => ({
    id: `s${i + 1}`,
    description: `Build object ${i + 1} for: ${prompt.slice(0, 40)}`,
    targetSheet: 'Dashboard',
    dependsOn: [] as string[],
    estimatedActions: 2,
  }));
  return {
    subtasks,
    clarificationsNeeded: [],
    confidence: 'high',
    reasoning: 'ok',
  };
}

describe('Spec 16 fix #3 — fallback-plan path unreachable under the scaled budget', () => {
  it.each(TIER3_4_BENCHMARK_PROMPTS)(
    'produces a real plan on the first attempt for: "%s"',
    async (prompt) => {
      const plan = planFor(prompt);
      const complete = jest.fn().mockImplementation((opts: { outcome?: { truncated?: boolean } }) => {
        // Simulates a well-behaved provider: given the (now complexity-scaled)
        // budget this call resolved, the model comfortably emits the full plan
        // without hitting the token cap.
        if (opts.outcome) opts.outcome.truncated = false;
        return Promise.resolve(JSON.stringify(plan));
      });
      const agent = buildAgent(complete);

      const result = await agent.plan(
        prompt,
        nonEmptyContext(),
        [],
        undefined,
        `corr_${prompt.slice(0, 8)}`,
        undefined,
        3,
      );

      // Fallback/last-resort machinery never engaged: exactly one call, and the
      // returned plan is the real one, not a raw-message stub.
      expect(complete).toHaveBeenCalledTimes(1);
      expect(result.subtasks.length).toBe(plan.subtasks.length);
      expect(result.reasoning).not.toContain('Fallback');
      expect(result.subtasks[0]?.description).not.toBe(prompt);
    },
  );

  it('never throws PlannerExhaustedError when every attempt is well-formed for a realistic compound prompt', async () => {
    const prompt = TIER3_4_BENCHMARK_PROMPTS[0]!;
    const plan = planFor(prompt);
    const complete = jest.fn().mockImplementation((opts: { outcome?: { truncated?: boolean } }) => {
      if (opts.outcome) opts.outcome.truncated = false;
      return Promise.resolve(JSON.stringify(plan));
    });
    const agent = buildAgent(complete);

    await expect(
      agent.plan(prompt, nonEmptyContext(), [], undefined, 'corr_no_exhaust', undefined, 3),
    ).resolves.not.toThrow();
  });

  it('scales the resolved budget by complexity score, not a flat per-tier value', () => {
    const simple = resolvePlannerMaxTokens(3, 'build a dashboard');
    const compound = resolvePlannerMaxTokens(3, TIER3_4_BENCHMARK_PROMPTS[0]);

    // A single-object Tier-3 request gets less than the old flat ceiling ...
    expect(simple).toBeLessThan(PLANNER_LAST_RESORT_MAX_TOKENS);
    // ... while the genuinely compound repro still reaches it, preserving the
    // headroom TASKS.md #170 established was necessary for that exact request.
    expect(compound).toBe(PLANNER_LAST_RESORT_MAX_TOKENS);
    expect(compound).toBeGreaterThan(simple);
  });

  it('gives every benchmark prompt a budget that leaves headroom beyond the reasoning cap', () => {
    for (const prompt of TIER3_4_BENCHMARK_PROMPTS) {
      const budget = resolvePlannerMaxTokens(3, prompt);
      // The Planner always requests a 1024-token reasoning cap (PLANNER_REASONING_MAX_TOKENS);
      // a budget that doesn't clear it plus real content room reproduces the spec 16 failure.
      expect(budget).toBeGreaterThan(1024 + 512);
    }
  });

  /**
   * Regression: a follow-up review asked what happens to a genuinely compound
   * request that happens to match zero TIER3_OBJECT_KEYWORDS — this prompt was
   * that case. Before the keyword list was widened it scored 1 (no hit at all:
   * "trends", "region", "month", "totals" weren't tracked) and got only the
   * 4096 base budget, risking the same wasted-retry-then-8192 round trip spec
   * 16 exists to avoid, despite clearly asking for a multi-cut breakdown.
   */
  it('scores "trends by region and by month, plus totals" above 1 after widening the keyword list', () => {
    const prompt = 'build me something showing trends by region and by month, plus totals';
    const score = resolveTier3ComplexityScore(prompt);
    expect(score).toBeGreaterThan(1);

    // And the resolved budget must actually reflect that — not just the score.
    const baseBudget = resolvePlannerMaxTokens(3, ''); // zero-keyword baseline
    const budget = resolvePlannerMaxTokens(3, prompt);
    expect(budget).toBeGreaterThan(baseBudget);
  });

  /**
   * Regression: "dashboard" doubles as a target-sheet name, so a single-object
   * request that merely names Dashboard as WHERE the object goes was over-
   * counted as two objects (found during a verification pass tracing
   * `resolveTier3ComplexityScore("Create a chart on Dashboard")`, which
   * returned 2 — chart AND dashboard — for a request that only asks for one
   * thing). `AMBIGUOUS_DASHBOARD_TARGET_SHEET_REF` now drops "dashboard" (and
   * any other TIER3_OBJECT_KEYWORDS word consumed inside the same
   * prepositional phrase, e.g. "sheet" in "on the Dashboard sheet") when it
   * only names a target sheet — but NOT when it's followed directly by a verb
   * rather than a clause boundary, which is what distinguishes these two
   * fixtures from the spec 16 repro just below.
   */
  it('does not double-count "dashboard" as a second object when it is only naming the target sheet', () => {
    expect(resolveTier3ComplexityScore('chart on Dashboard')).toBe(1);
    expect(resolveTier3ComplexityScore('Create a chart on Dashboard')).toBe(1);
    expect(resolvePlannerMaxTokens(3, 'Create a chart on Dashboard')).toBe(4096);

    // "sheet" is independently a TIER3_OBJECT_KEYWORDS word — must not survive
    // as a leftover second object once it's recognized as part of the SAME
    // "the Dashboard sheet" target-sheet reference as "dashboard" itself.
    expect(resolveTier3ComplexityScore('put a summary on the Dashboard sheet')).toBe(1);
    expect(resolvePlannerMaxTokens(3, 'put a summary on the Dashboard sheet')).toBe(4096);
  });

  it('still counts "dashboard" as a genuine object in the spec 16 repro, where it is not a bare target-sheet reference', () => {
    const repro = TIER3_4_BENCHMARK_PROMPTS[0]!; // "In dashboard create a chart ,and analysis ..."
    // "In dashboard create a chart" — "dashboard" is followed by the verb
    // "create", not a clause boundary, so it is NOT treated as ambiguous.
    const score = resolveTier3ComplexityScore(repro);
    expect(score).toBeGreaterThanOrEqual(4);
    expect(resolvePlannerMaxTokens(3, repro)).toBe(PLANNER_LAST_RESORT_MAX_TOKENS);
  });
});
