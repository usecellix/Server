import { summarizePlanIntent, buildUserFacingSummary } from '../src/excel-ai/utils/user-facing-response.util';

/**
 * TASKS.md #155 — the Accept card promises only what was actually delivered.
 *
 * #149 made the card describe the PLAN rather than the emitted actions, which
 * is far more readable — and introduced a way to over-promise. A live smoke run
 * caught exactly that: the card said "Write Consolidated Transactions header at
 * Main!A18" for a run whose actions never touched row 18, and listed a month
 * sheet whose Executor emitted only ADD_SHEET and no headers.
 *
 * This is the `CODEBASE_ANALYSIS.md` §3.7 shape again — the system knew the
 * work was incomplete (it had both the plan and the emitted actions) and said
 * nothing.
 */
describe('plan/delivery gap — the card must not promise undelivered work', () => {
  const planned = [
    { id: 's1', description: "Create sheet 'Main' if it doesn't exist", targetSheet: 'Main' },
    { id: 's2', description: 'Write Consolidated Transactions header at Main!A18', targetSheet: 'Main' },
    { id: 'm4', description: "Create sheet 'April' and set A1:J1 headers", targetSheet: 'April' },
  ];

  /** Mirrors the orchestrator's filter: delivered = produced >=1 action. */
  function delivered(ids: string[]) {
    return planned.filter((p) => ids.includes(p.id));
  }

  it('excludes a subtask that emitted no actions', () => {
    // s2 and m4 planned but never executed — the observed live failure.
    const bullets = summarizePlanIntent(delivered(['s1']));
    expect(bullets.join(' ')).not.toContain('Consolidated Transactions');
    expect(bullets.join(' ')).not.toContain('April');
  });

  it('keeps subtasks that did deliver', () => {
    const bullets = summarizePlanIntent(delivered(['s1', 's2']));
    expect(bullets.join(' ')).toContain('Consolidated Transactions');
  });

  it('a fully delivered plan is described in full', () => {
    const bullets = summarizePlanIntent(delivered(['s1', 's2', 'm4']));
    expect(bullets).toHaveLength(3);
  });

  it('the card falls back to the action rollup when NOTHING was delivered', () => {
    const summary = buildUserFacingSummary({
      actions: [
        { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'x' },
        { type: 'SORT_RANGE', sheetName: 'Main', range: 'A1:C9' },
      ],
      changes: [],
      planSubtasks: [],
    });
    // Never silent: with no intent to show it must still describe the actions.
    expect(summary.bullets).toBeDefined();
    expect(summary.bullets!.length).toBeGreaterThan(0);
  });
});
