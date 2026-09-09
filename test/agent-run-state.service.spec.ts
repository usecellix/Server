import { AgentRunStateService } from '../src/agents/agent-run-state.service';
import { AgentRunDocument } from '../src/agents/schemas/agent-run.schema';
import { SubTask } from '../src/agents/types/agent.types';

/**
 * STEPWISE_EXECUTION.md SD-4 — a rejected or skipped wave cascade-skips every
 * subtask that transitively depends on it, and the run continues rather than
 * aborting. Getting the cascade wrong is what would let "populate Main" run
 * against a Main sheet whose creation the user just rejected.
 */

const subtask = (id: string, dependsOn: string[] = []): SubTask => ({
  id,
  description: `subtask ${id}`,
  targetSheet: 'Main',
  dependsOn,
  estimatedActions: 1,
});

/** Minimal stand-in for a Mongoose document — save/markModified are no-ops. */
function makeRun(subtasks: SubTask[], waves: string[][], waveIndex: number): AgentRunDocument {
  return {
    runId: 'run_test',
    conversationId: 'conv_1',
    traceId: 'trace_1',
    prompt: 'build it',
    status: 'awaiting_decision',
    waveIndex,
    waveTotal: waves.length,
    subtasks,
    waves,
    subtaskStates: subtasks.map((entry) => ({
      subtaskId: entry.id,
      actions: [],
      completed: false,
    })),
    context: { activeSheetName: 'Main', sheets: [], namedRanges: [], tables: [] },
    changeSetIds: [],
    conversationHistory: [],
    markModified: () => undefined,
    save: async () => undefined,
  } as unknown as AgentRunDocument;
}

describe('AgentRunStateService — decisions and cascade skipping', () => {
  const service = new AgentRunStateService({} as never);

  it('accepting a wave skips nothing downstream', async () => {
    const subtasks = [subtask('s1'), subtask('s2', ['s1'])];
    const run = makeRun(subtasks, [['s1'], ['s2']], 0);

    const { cascadeSkipped } = await service.applyDecision(run, 'accepted');

    expect(cascadeSkipped).toEqual([]);
    expect(run.subtaskStates.find((s) => s.subtaskId === 's2')?.decision).toBeUndefined();
  });

  it('rejecting a wave cascade-skips its transitive dependents', async () => {
    // s3 depends on s2 depends on s1 — rejecting s1 must reach s3, not just s2.
    const subtasks = [subtask('s1'), subtask('s2', ['s1']), subtask('s3', ['s2'])];
    const run = makeRun(subtasks, [['s1'], ['s2'], ['s3']], 0);

    const { cascadeSkipped } = await service.applyDecision(run, 'rejected');

    expect(cascadeSkipped.sort()).toEqual(['s2', 's3']);
    expect(run.subtaskStates.find((s) => s.subtaskId === 's3')?.decision).toBe('skipped');
  });

  it('leaves independent subtasks alone when a sibling is rejected', async () => {
    const subtasks = [subtask('s1'), subtask('s2', ['s1']), subtask('independent')];
    const run = makeRun(subtasks, [['s1', 'independent'], ['s2']], 0);
    // Only s1/independent are in wave 0; rejecting that wave rejects both.
    const { cascadeSkipped } = await service.applyDecision(run, 'rejected');

    // s2 depends on s1 so it cascades; nothing else exists to spare here, but
    // the decision must not invent skips for subtasks with no such edge.
    expect(cascadeSkipped).toEqual(['s2']);
  });

  it('nextExecutableWave skips waves whose subtasks are all already decided', () => {
    const subtasks = [subtask('s1'), subtask('s2', ['s1']), subtask('s3', ['s2'])];
    const run = makeRun(subtasks, [['s1'], ['s2'], ['s3']], 0);
    run.subtaskStates.find((s) => s.subtaskId === 's2')!.decision = 'skipped';

    const next = service.nextExecutableWave(run);

    // Wave 1 is fully decided (cascade-skipped), so the next real work is wave 2.
    expect(next?.waveIndex).toBe(2);
    expect(next?.subtasks.map((s) => s.id)).toEqual(['s3']);
  });

  it('returns null once every remaining wave is decided', () => {
    const subtasks = [subtask('s1'), subtask('s2', ['s1'])];
    const run = makeRun(subtasks, [['s1'], ['s2']], 0);
    run.subtaskStates.find((s) => s.subtaskId === 's2')!.decision = 'skipped';

    expect(service.nextExecutableWave(run)).toBeNull();
  });

  it('an empty wave marked skipped always advances, so the retry loop terminates', async () => {
    // The empty-wave path in executeStepwiseWave recurses: mark skipped, try the
    // next wave. If applyDecision failed to actually decide the current wave,
    // nextExecutableWave would hand back the SAME wave forever and the request
    // would spin until the process died. This pins the termination guarantee.
    const subtasks = [subtask('s1'), subtask('s2'), subtask('s3')];
    const run = makeRun(subtasks, [['s1'], ['s2'], ['s3']], 0);

    const before = service.nextExecutableWave(run);
    expect(before?.waveIndex).toBe(1);

    // Simulate the empty-wave branch for wave 0.
    await service.applyDecision(run, 'skipped');
    expect(run.subtaskStates.find((s) => s.subtaskId === 's1')?.decision).toBe('skipped');

    // Wave 0 is now decided, so the next call cannot return it again.
    const after = service.nextExecutableWave(run);
    expect(after?.waveIndex).toBe(1);
    expect(after?.subtasks.map((s) => s.id)).toEqual(['s2']);
  });

  describe('applyReadback — rejecting the frontend WorkbookContext shape', () => {
    // Live incident (Sept 8, 2026): the frontend's WorkbookContext.sheets is
    // `SheetSnapshot[]` (types/cellix.types.ts — `sheetName`, `colCount`,
    // `headers`, `sampleData`), not this backend's `SheetContext[]` (`name`,
    // `values`, `formulas`, `numberFormats`). Sending the frontend shape
    // through unvalidated appended a malformed sheet object into
    // `run.context.sheets`, and the next wave's Executor crashed reading
    // `.values.length` on it — a live "Cannot read properties of undefined
    // (reading 'length')" failure with zero context, since it propagated as a
    // bare, unwrapped Node TypeError all the way to the SSE `error` event.

    const realSheet = {
      name: 'Main',
      usedRange: 'A1:D4',
      rowCount: 4,
      columnCount: 4,
      values: [['Month', 'Total', 'Paid', 'Pending']],
      formulas: [['', '', '', '']],
      numberFormats: [['General', 'General', 'General', 'General']],
      structure: 'unknown' as const,
      headerRowIndex: 0,
    };

    it('leaves context.sheets untouched when every readback entry is the mismatched frontend shape', async () => {
      const run = makeRun([subtask('s1')], [['s1']], -1);
      run.context = { activeSheetName: 'Main', sheets: [realSheet], namedRanges: [], tables: [] };

      // The exact SheetSnapshot shape the frontend actually sends — no `name`,
      // no `values`/`formulas`/`numberFormats` arrays.
      const frontendShapedReadback = [
        { sheetName: 'Main', usedRange: 'A1:D4', rowCount: 4, colCount: 4, headers: ['Month'], sampleData: [] },
      ];

      await service.applyReadback(run, frontendShapedReadback as never);

      expect(run.context.sheets).toEqual([realSheet]);
      expect(run.context.sheets).toHaveLength(1);
      // The bug appended a SECOND, malformed entry — confirm that never happens.
      expect(run.context.sheets.every((s) => Array.isArray(s.values))).toBe(true);
    });

    it('accepts a readback entry that genuinely matches SheetContext', async () => {
      const run = makeRun([subtask('s1')], [['s1']], -1);
      run.context = { activeSheetName: 'Main', sheets: [realSheet], namedRanges: [], tables: [] };

      const updatedSheet = { ...realSheet, values: [['Month', 'Total', 'Paid', 'Pending'], ['January', 100, 80, 20]] };

      await service.applyReadback(run, [updatedSheet]);

      expect(run.context.sheets).toEqual([updatedSheet]);
    });

    it('ignores an empty or absent readback without touching context', async () => {
      const run = makeRun([subtask('s1')], [['s1']], -1);
      run.context = { activeSheetName: 'Main', sheets: [realSheet], namedRanges: [], tables: [] };

      await service.applyReadback(run, undefined);
      expect(run.context.sheets).toEqual([realSheet]);

      await service.applyReadback(run, []);
      expect(run.context.sheets).toEqual([realSheet]);
    });
  });

  it('summarizes skipped subtasks for an honest closing summary', () => {
    const subtasks = [subtask('s1'), subtask('s2', ['s1'])];
    const run = makeRun(subtasks, [['s1'], ['s2']], 0);
    run.subtaskStates.find((s) => s.subtaskId === 's2')!.decision = 'skipped';

    const skipped = service.summarizeSkipped(run);

    expect(skipped).toHaveLength(1);
    expect(skipped[0].subtaskId).toBe('s2');
    expect(skipped[0].description).toBe('subtask s2');
  });
});
