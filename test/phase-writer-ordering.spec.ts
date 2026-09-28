import { ensureWritersDependOnCreator } from '../src/agents/utils/plan-coverage.util';
import { computeExecutionWaves } from '../src/agents/utils/task-graph.util';
import { PlannerOutput, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #299 — a sheet written to before it is created.
 *
 * Live shape, from the smoke run after #296: Main is built by several
 * subtasks. `p3_s4` carries the ADD_SHEET and emits it first, correctly. The
 * others only write. Nothing forced them to depend on the creator, and
 * `computeExecutionWaves` orders purely on dependsOn edges, so a writer landed
 * in the same wave as the creator and its SET_CELL / FORMAT_RANGE / BATCH_SET
 * / SET_COLUMN_WIDTH went out ahead of the create. In Excel that is "The
 * requested resource doesn't exist" at Accept.
 */

const context: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [{
    name: 'Sheet1', usedRange: 'A1:A1', rowCount: 1, columnCount: 1,
    values: [['']], formulas: [['']], numberFormats: [['General']],
    structure: 'data_table', headerRowIndex: 0,
  }],
  namedRanges: [],
  tables: [],
};

const sub = (over: Partial<SubTask> & { id: string }): SubTask => ({
  targetSheet: 'Main',
  dependsOn: [],
  estimatedActions: 5,
  description: 'Write KPI formulas on Main',
  ...over,
});

const planOf = (subtasks: SubTask[]): PlannerOutput => ({
  subtasks,
  clarificationsNeeded: [],
  confidence: 'high',
  reasoning: '',
});

describe('ensureWritersDependOnCreator (TASKS.md #299)', () => {
  const creator = sub({
    id: 'p3_s4',
    description: "Create sheet 'Main' at position 0 and write the dashboard title.",
  });

  it('orders every Main writer after the subtask that creates Main', () => {
    const plan = planOf([
      creator,
      sub({ id: 'p3_s5' }),
      sub({ id: 'p3_s6', description: 'Add charts to Main' }),
    ]);

    const { plan: fixed, added } = ensureWritersDependOnCreator(plan, context);

    expect(added.map((e) => e.writer).sort()).toEqual(['p3_s5', 'p3_s6']);
    for (const id of ['p3_s5', 'p3_s6']) {
      expect(fixed.subtasks.find((s) => s.id === id)?.dependsOn).toContain('p3_s4');
    }
  });

  it('the creator ends up in an EARLIER wave than its writers', () => {
    // The assertion that actually matters: the edge has to change scheduling,
    // not just the plan document.
    const plan = planOf([creator, sub({ id: 'p3_s5' }), sub({ id: 'p3_s6' })]);

    const before = computeExecutionWaves(plan.subtasks);
    const beforeWave = (id: string) => before.findIndex((w) => w.some((s) => s.id === id));
    expect(beforeWave('p3_s5')).toBe(beforeWave('p3_s4')); // the bug: same wave

    const after = computeExecutionWaves(ensureWritersDependOnCreator(plan, context).plan.subtasks);
    const afterWave = (id: string) => after.findIndex((w) => w.some((s) => s.id === id));
    expect(afterWave('p3_s4')).toBeLessThan(afterWave('p3_s5'));
    expect(afterWave('p3_s4')).toBeLessThan(afterWave('p3_s6'));
  });

  it('adds nothing when the writers already depend on the creator', () => {
    const plan = planOf([creator, sub({ id: 'p3_s5', dependsOn: ['p3_s4'] })]);
    expect(ensureWritersDependOnCreator(plan, context).added).toEqual([]);
  });

  it('adds nothing when the ordering is already indirect', () => {
    // p3_s6 -> p3_s5 -> p3_s4 is already correct; a direct edge would be noise.
    const plan = planOf([
      creator,
      sub({ id: 'p3_s5', dependsOn: ['p3_s4'] }),
      sub({ id: 'p3_s6', dependsOn: ['p3_s5'] }),
    ]);
    expect(ensureWritersDependOnCreator(plan, context).added).toEqual([]);
  });

  it('never adds an edge that would close a cycle', () => {
    // The creator itself depends on the writer. Adding writer -> creator would
    // strand the run; leaving it unordered merely risks the original bug.
    const plan = planOf([
      sub({ id: 'p3_s4', description: "Create sheet 'Main'.", dependsOn: ['p3_s5'] }),
      sub({ id: 'p3_s5' }),
    ]);
    const { added, plan: fixed } = ensureWritersDependOnCreator(plan, context);

    expect(added).toEqual([]);
    expect(() => computeExecutionWaves(fixed.subtasks)).not.toThrow();
  });

  it('leaves a sheet that already exists in the workbook alone', () => {
    // Nothing needs creating, so nothing needs ordering.
    const plan = planOf([
      sub({ id: 'a', targetSheet: 'Sheet1', description: "Create sheet 'Sheet1'." }),
      sub({ id: 'b', targetSheet: 'Sheet1' }),
    ]);
    expect(ensureWritersDependOnCreator(plan, context).added).toEqual([]);
  });

  it('does not order month sheets against each other', () => {
    // Twelve independent creators must stay parallel — this net must never
    // serialise a wave that is legitimately wide.
    const months = ['January', 'February', 'March'];
    const plan = planOf(
      months.map((m) => sub({ id: `p2_${m}`, targetSheet: m, description: `Create sheet '${m}'.` })),
    );
    const { added, plan: fixed } = ensureWritersDependOnCreator(plan, context);

    expect(added).toEqual([]);
    expect(computeExecutionWaves(fixed.subtasks)).toHaveLength(1);
  });
});
