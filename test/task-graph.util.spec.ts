import { computeExecutionWaves } from '../src/agents/utils/task-graph.util';
import { SubTask } from '../src/agents/types/agent.types';

describe('computeExecutionWaves', () => {
  it('groups independent subtasks into the same wave', () => {
    const subtasks: SubTask[] = [
      {
        id: 's1',
        description: 'Format Summary',
        targetSheet: 'Summary',
        dependsOn: [],
        estimatedActions: 1,
      },
      {
        id: 's2',
        description: 'Format Data',
        targetSheet: 'Data',
        dependsOn: [],
        estimatedActions: 1,
      },
      {
        id: 's3',
        description: 'Update totals',
        targetSheet: 'Summary',
        dependsOn: ['s1'],
        estimatedActions: 1,
      },
    ];

    const waves = computeExecutionWaves(subtasks);

    expect(waves).toHaveLength(2);
    expect(waves[0].map((subtask) => subtask.id).sort()).toEqual(['s1', 's2']);
    expect(waves[1].map((subtask) => subtask.id)).toEqual(['s3']);
  });
});

/**
 * TASKS.md #83: the planner prompt now emits Main-sheet subtasks BEFORE the 12
 * month-sheet subtasks, so that a token-budget truncation loses cheap, repetitive
 * boilerplate rather than the dashboard the user actually asked for.
 *
 * That reordering is only safe because execution order comes from `dependsOn`,
 * never from array position. These tests pin that invariant — if scheduling ever
 * became order-sensitive, the reordered prompt would start creating Main formulas
 * that reference month sheets which do not exist yet.
 */
describe('computeExecutionWaves — order independence (TASKS.md #83)', () => {
  function ledgerPlan(mainFirst: boolean): SubTask[] {
    const months = ['January', 'February', 'March'].map((name, i) => ({
      id: `m${i + 1}`,
      description: `Create sheet '${name}' and write header row`,
      targetSheet: name,
      dependsOn: [] as string[],
      estimatedActions: 1,
    }));
    const main: SubTask[] = [
      {
        id: 'main_create',
        description: "Create sheet 'Main'",
        targetSheet: 'Main',
        dependsOn: [],
        estimatedActions: 1,
      },
      {
        id: 'main_totals',
        description: 'Monthly Totals table referencing each month sheet',
        targetSheet: 'Main',
        // Explicit edges are what make the reordering safe.
        dependsOn: ['m1', 'm2', 'm3', 'main_create'],
        estimatedActions: 18,
      },
      {
        id: 'main_consolidated',
        description: 'Consolidated Transactions header at Main!A18',
        targetSheet: 'Main',
        dependsOn: ['main_totals'],
        estimatedActions: 1,
      },
    ];
    return mainFirst ? [...main, ...months] : [...months, ...main];
  }

  function waveIndexOf(waves: SubTask[][], id: string): number {
    return waves.findIndex((wave) => wave.some((s) => s.id === id));
  }

  it('schedules month sheets before the Main formulas that reference them, even when Main is emitted first', () => {
    const waves = computeExecutionWaves(ledgerPlan(true));

    const monthWave = Math.max(
      waveIndexOf(waves, 'm1'),
      waveIndexOf(waves, 'm2'),
      waveIndexOf(waves, 'm3'),
    );
    // The dependency edge, not the array position, decides this.
    expect(waveIndexOf(waves, 'main_totals')).toBeGreaterThan(monthWave);
    expect(waveIndexOf(waves, 'main_consolidated')).toBeGreaterThan(
      waveIndexOf(waves, 'main_totals'),
    );
  });

  it('produces the same wave assignment whichever order the planner emitted subtasks in', () => {
    const assign = (plan: SubTask[]) => {
      const waves = computeExecutionWaves(plan);
      return Object.fromEntries(
        plan.map((s) => [s.id, waveIndexOf(waves, s.id)]),
      );
    };

    expect(assign(ledgerPlan(true))).toEqual(assign(ledgerPlan(false)));
  });

  it('never strands a subtask when Main is emitted first', () => {
    const plan = ledgerPlan(true);
    const waves = computeExecutionWaves(plan);
    expect(waves.flat()).toHaveLength(plan.length);
    // Every subtask lands in a real wave (no -1 / stranded bucket surprises).
    for (const subtask of plan) {
      expect(waveIndexOf(waves, subtask.id)).toBeGreaterThanOrEqual(0);
    }
  });
});
