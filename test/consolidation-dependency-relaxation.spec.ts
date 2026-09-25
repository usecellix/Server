import {
  relaxConsolidationDependencies,
  splitSpecPinnedSubtasks,
} from '../src/agents/utils/header-table-split.util';
import { computeExecutionWaves } from '../src/agents/utils/task-graph.util';
import { SubTask } from '../src/agents/types/agent.types';

/**
 * TASKS.md #309 — a dashboard must not be lost because one month sheet's
 * optional tail failed.
 *
 * The most expensive failure shape in LONG_PROMPT_RELIABILITY_PLAN.md's live
 * scoreboard, twice over: a run builds all twelve month sheets correctly, one
 * (or all) of the month FORMULA steps fails to converge, and Main — which
 * depends on those steps — is gated off and never built. The user gets twelve
 * good sheets and no dashboard.
 *
 * After Phase 1.5, the part Main actually needs (sheet + header row + table)
 * belongs to the deterministic header step, which makes no model call and so
 * cannot time out, drift or hit an iteration cap. Pointing Main at that step
 * is both truer to what it needs and far more robust.
 */

const COLUMNS = ['Unit No', 'Guest', 'Check In', 'Check Out', 'Total Amount'];

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const monthSubtask = (month: string, i: number): SubTask => ({
  id: `p2_s${i + 1}`,
  targetSheet: month,
  dependsOn: ['p1_s1'],
  estimatedActions: 26,
  expectedHeaders: COLUMNS,
  description:
    `Create sheet '${month}' with headers in row 1 (A1:E1): ${COLUMNS.join(', ')}. ` +
    `Create table 'tbl${month}' over A1:E2. Set row-2 formulas and column widths.`,
});

/** The live plan shape: Lists, twelve months, and a Main that reads them all. */
function livePlan(): SubTask[] {
  return [
    {
      id: 'p1_s1',
      targetSheet: 'Lists',
      dependsOn: [],
      estimatedActions: 4,
      description: "Create sheet 'Lists' with the dropdown values.",
    },
    ...MONTHS.map(monthSubtask),
    {
      id: 'p3_s1',
      targetSheet: 'Main',
      dependsOn: MONTHS.map((_, i) => `p2_s${i + 1}`),
      estimatedActions: 20,
      description: "Create sheet 'Main' with a dashboard consolidating every month.",
    },
  ];
}

const waveOf = (waves: SubTask[][], id: string) =>
  waves.findIndex((wave) => wave.some((s) => s.id === id));

describe('relaxConsolidationDependencies (TASKS.md #309)', () => {
  it('points the dashboard at the steps that CREATE the sheets, not the ones that finish them', () => {
    const split = splitSpecPinnedSubtasks(livePlan());
    const { subtasks, relaxed } = relaxConsolidationDependencies(split);

    const main = subtasks.find((s) => s.id === 'p3_s1');
    expect(main?.dependsOn.sort()).toEqual(MONTHS.map((m) => `hdr_${m}`).sort());
    expect(relaxed).toHaveLength(MONTHS.length);
  });

  it('the dashboard is still built when a month\'s formula step fails', () => {
    // The whole point. `p2_s4` (April) failing must not remove Main from the
    // graph that reaches it.
    const split = splitSpecPinnedSubtasks(livePlan());
    const { subtasks } = relaxConsolidationDependencies(split);

    const main = subtasks.find((s) => s.id === 'p3_s1');
    expect(main?.dependsOn).not.toContain('p2_s4');
    // April's own sheet still has to exist first.
    expect(main?.dependsOn).toContain('hdr_April');
  });

  it('survives the run-3 shape where EVERY month tail failed', () => {
    // All twelve rest steps failing took Main down with them and the run
    // ended at wave 4 of 6. Main now depends on nothing that can fail that
    // way: every one of its dependencies is a zero-LLM-call step.
    const split = splitSpecPinnedSubtasks(livePlan());
    const { subtasks } = relaxConsolidationDependencies(split);

    const main = subtasks.find((s) => s.id === 'p3_s1');
    const deterministicIds = new Set(
      subtasks.filter((s) => s.isDeterministicHeaderStep).map((s) => s.id),
    );
    expect(main?.dependsOn.every((dep) => deterministicIds.has(dep))).toBe(true);
  });

  it('a rest step keeps the dependency on its OWN header step', () => {
    // That one is real sequencing, not consolidation: the sheet has to exist
    // before its formulas can be written into it.
    const split = splitSpecPinnedSubtasks(livePlan());
    const { subtasks } = relaxConsolidationDependencies(split);

    const april = subtasks.find((s) => s.id === 'p2_s4');
    expect(april?.dependsOn).toContain('hdr_April');
  });

  it('schedules the dashboard a wave earlier, alongside the month formulas', () => {
    const split = splitSpecPinnedSubtasks(livePlan());
    const before = computeExecutionWaves(split);
    const after = computeExecutionWaves(relaxConsolidationDependencies(split).subtasks);

    // Before: Main waits for the whole formula wave. After: it only waits for
    // the header wave, so it runs in parallel with the formulas instead.
    expect(waveOf(before, 'p3_s1')).toBeGreaterThan(waveOf(before, 'p2_s1'));
    expect(waveOf(after, 'p3_s1')).toBe(waveOf(after, 'p2_s1'));
  });

  it('leaves a plan with no deterministic header steps completely alone', () => {
    const plain: SubTask[] = [
      { id: 'a', targetSheet: 'Sheet1', dependsOn: [], estimatedActions: 2, description: 'x' },
      { id: 'b', targetSheet: 'Sheet2', dependsOn: ['a'], estimatedActions: 2, description: 'y' },
    ];
    const { subtasks, relaxed } = relaxConsolidationDependencies(plain);

    expect(relaxed).toEqual([]);
    expect(subtasks).toBe(plain);
  });

  it('does not disturb a dependency that has nothing to do with the split', () => {
    // Main also depends on Lists, which was never split. That edge must survive.
    const plan = livePlan();
    const main = plan.find((s) => s.id === 'p3_s1')!;
    main.dependsOn = ['p1_s1', ...main.dependsOn];

    const { subtasks } = relaxConsolidationDependencies(splitSpecPinnedSubtasks(plan));
    expect(subtasks.find((s) => s.id === 'p3_s1')?.dependsOn).toContain('p1_s1');
  });

  it('produces a graph with no cycles and no dangling ids', () => {
    // A relaxation that stranded the run would be far worse than the bug.
    const split = splitSpecPinnedSubtasks(livePlan());
    const { subtasks } = relaxConsolidationDependencies(split);

    const ids = new Set(subtasks.map((s) => s.id));
    for (const subtask of subtasks) {
      for (const dep of subtask.dependsOn) expect(ids.has(dep)).toBe(true);
    }

    const waves = computeExecutionWaves(subtasks);
    expect(waves.flat()).toHaveLength(subtasks.length);
  });
});
