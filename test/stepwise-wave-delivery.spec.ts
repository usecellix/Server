import { AgenticLoopService } from '../src/agents/agenticLoop.service';
import { CompletenessChecker } from '../src/agents/checkers/completeness.checker';
import { FormattingChecker } from '../src/agents/checkers/formatting.checker';
import { OverwriteOccupancyChecker } from '../src/agents/checkers/overwrite-occupancy.checker';
import { Action, SubTask } from '../src/agents/types/agent.types';
import { dedupeIdenticalWrites } from '../src/agents/utils/dedupe-writes.util';

/**
 * Two live runs on Sept 25, 2026, one root cause each for three symptoms.
 *
 * #315 — a stepwise wave holds only ITS OWN subtask states, so a dependency on
 * an earlier wave's subtask was invisible to `filterSafePartialDelivery` and
 * read as unmet. The moment any peer in the wave failed, every finished sibling
 * depending on an earlier wave was dropped with no reason: Main's Monthly
 * Totals in one run, all 12 month sheets in the other.
 *
 * #318 — in the second run those 12 months still SHIPPED (a different branch
 * flattens `state.completed`) while being recorded as not done.
 *
 * #317 — eleven parallel month subtasks each wrote the identical Lists cells;
 * the client refused the second copy as an overwrite, so the step could never
 * be accepted.
 */

function makeService(): AgenticLoopService {
  return new AgenticLoopService(
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    new CompletenessChecker(),
    new FormattingChecker(),
    new OverwriteOccupancyChecker(),
  );
}

const subtask = (id: string, dependsOn: string[] = []): SubTask => ({
  id,
  description: `subtask ${id}`,
  targetSheet: 'Main',
  dependsOn,
  estimatedActions: 1,
});

const state = (
  task: SubTask,
  actions: Action[],
  extra: { completed?: boolean; failedReason?: string; verified?: boolean } = {},
) => ({ subtask: task, actions, droppedActions: [], completed: true, ...extra });

const cell = (sheetName: string, address: string, value: string) =>
  ({ type: 'SET_CELL', sheetName, address, value }) as unknown as Action;

type LoopResult = {
  actions: Action[];
  completedSubtasks: Array<{ subtaskId: string }>;
  failedSubtasks: Array<{ subtaskId: string; reason: string }>;
};

function build(states: unknown[]): LoopResult {
  return (makeService() as any).buildLoopResult(states, 1, false, {
    preferCompletedOnly: true,
    defaultFailReason: 'Could not complete and verify the full request',
  });
}

describe('stepwise wave delivery — earlier-wave dependencies (TASKS.md #315)', () => {
  it('delivers a finished sibling whose dependency is from an earlier wave, when a peer fails', () => {
    // Run 2's shape: p3_s2 depends on p3_s1 and hdr_* (all earlier waves);
    // p3_s5 in the same wave failed.
    const totals = subtask('p3_s2', ['p3_s1', 'hdr_January']);
    const header = subtask('p3_s5', ['p3_s1']);
    const result = build([
      state(totals, [cell('Main', 'A5', 'January')]),
      state(header, [], { completed: false, failedReason: 'Could not complete after 2 attempts' }),
    ]);

    expect(result.completedSubtasks.map((s) => s.subtaskId)).toEqual(['p3_s2']);
    expect(result.actions).toHaveLength(1);
    expect(result.failedSubtasks.map((f) => f.subtaskId)).toEqual(['p3_s5']);
  });

  it('still holds back a subtask whose dependency failed IN THIS wave — and says why', () => {
    const base = subtask('a');
    const dependent = subtask('b', ['a']);
    const result = build([
      state(base, [], { completed: false, failedReason: 'hit max iterations (10)' }),
      state(dependent, [cell('Main', 'B2', 'x')]),
    ]);

    expect(result.completedSubtasks).toHaveLength(0);
    expect(result.actions).toHaveLength(0);
    const held = result.failedSubtasks.find((f) => f.subtaskId === 'b');
    expect(held?.reason).toMatch(/depends on/);
  });
});

describe('stepwise wave delivery — recorded equals shipped (TASKS.md #318)', () => {
  it('every subtask whose actions ship is reported as delivered', () => {
    const months = ['January', 'February', 'March'].map((m) => subtask(`p2_${m}`, [`hdr_${m}`]));
    const main = subtask('p3_s1');
    const result = build([
      ...months.map((m, i) => state(m, [cell(m.id, 'F2', `=E2-D2 ${i}`)])),
      state(main, [], { completed: false, failedReason: 'E20 contains an empty formula' }),
    ]);

    const shipped = new Set(
      (result.actions as unknown as Array<{ sheetName: string }>).map((a) => a.sheetName),
    );
    const recorded = new Set(result.completedSubtasks.map((s) => s.subtaskId));
    expect(recorded).toEqual(shipped);
    expect(recorded.size).toBe(3);
  });

  it('a finished subtask the verifier rejected, shipped by the fallback branch, is recorded as delivered', () => {
    // Nothing survives the partial-delivery filter, so the fallback branch
    // flattens `state.completed` — which ships this subtask's actions.
    const task = subtask('p3_s1');
    const result = build([
      state(task, [cell('Main', 'A1', 'Dashboard')], { failedReason: 'E20 contains an empty formula' }),
    ]);

    expect(result.actions).toHaveLength(1);
    expect(result.completedSubtasks.map((s) => s.subtaskId)).toEqual(['p3_s1']);
  });
});

describe('identical writes across one wave (TASKS.md #317)', () => {
  const listsWrite = () =>
    ({
      type: 'BATCH_SET',
      sheetName: 'Lists',
      operations: [
        { address: 'D1', value: 'Unit No' },
        { address: 'D2', value: '101' },
      ],
    }) as unknown as Action;

  it('keeps one copy of the same BATCH_SET emitted by several subtasks', () => {
    const { actions, removed, conflicts } = dedupeIdenticalWrites([listsWrite(), listsWrite(), listsWrite()]);
    expect(actions).toEqual([listsWrite()]);
    expect(removed).toBe(4);
    expect(conflicts).toEqual([]);
  });

  it('treats an address and its row/col form as the same cell', () => {
    const { actions } = dedupeIdenticalWrites([
      cell('Lists', 'D1', 'Unit No'),
      { type: 'SET_CELL', sheetName: 'Lists', row: 0, col: 3, value: 'Unit No' } as unknown as Action,
    ]);
    expect(actions).toHaveLength(1);
  });

  it('leaves DIFFERENT values for one cell in place and reports the conflict', () => {
    const { actions, removed, conflicts } = dedupeIdenticalWrites([
      cell('Lists', 'D1', 'Unit No'),
      cell('Lists', 'D1', 'Unit'),
    ]);
    expect(actions).toHaveLength(2);
    expect(removed).toBe(0);
    expect(conflicts).toEqual(['lists|0|3']);
  });

  it('is applied to what a wave ships', () => {
    const a = subtask('p2_January');
    const b = subtask('p2_February');
    const result = (makeService() as any).buildLoopResult(
      [state(a, [listsWrite()]), state(b, [listsWrite()])],
      1,
      true,
      { preferCompletedOnly: false, defaultFailReason: 'x' },
    ) as LoopResult;
    expect(result.actions).toHaveLength(1);
  });
});
