import { AgentRunStateService } from '../src/agents/agent-run-state.service';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 8 (TASKS.md #293).
 *
 * Observed live: a stepwise build finished its first wave, the task pane lost
 * its connection before calling /continue, and the run sat in
 * `awaiting_decision` with its applied waves stranded and no route back. The
 * only option was to rebuild from scratch. Long builds are hit hardest for the
 * obvious reason that they are long.
 */
describe('AgentRunStateService.findResumableRun (Phase 8)', () => {
  function buildService(run: unknown): AgentRunStateService {
    const model = {
      findOne: jest.fn().mockReturnValue({ sort: jest.fn().mockResolvedValue(run) }),
    };
    return new AgentRunStateService(model as never);
  }

  const runDoc = (overrides: Record<string, unknown> = {}) => ({
    runId: 'run_1',
    conversationId: 'conv_1',
    status: 'awaiting_decision',
    waveIndex: 0,
    waveTotal: 5,
    subtasks: [
      { id: 's1', description: 'a', targetSheet: 'January', dependsOn: [], estimatedActions: 1 },
      { id: 's2', description: 'b', targetSheet: 'February', dependsOn: [], estimatedActions: 1 },
    ],
    waves: [['s1'], ['s2']],
    subtaskStates: [
      { subtaskId: 's1', actions: [], completed: true, decision: 'accepted' },
      { subtaskId: 's2', actions: [], completed: false },
    ],
    ...overrides,
  });

  it('finds the stranded run so the panel can offer to carry on', async () => {
    const run = await buildService(runDoc()).findResumableRun('conv_1');
    expect(run?.runId).toBe('run_1');
  });

  it('returns nothing when there is no unfinished run at all', async () => {
    expect(await buildService(null).findResumableRun('conv_1')).toBeNull();
  });

  it('returns nothing when every wave is already decided — finished is not resumable', async () => {
    const done = runDoc({
      waveIndex: 1,
      subtaskStates: [
        { subtaskId: 's1', actions: [], completed: true, decision: 'accepted' },
        { subtaskId: 's2', actions: [], completed: true, decision: 'accepted' },
      ],
    });
    expect(await buildService(done).findResumableRun('conv_1')).toBeNull();
  });

  it('refuses a run belonging to someone else', async () => {
    const owned = runDoc({ userId: 'user_a' });
    expect(await buildService(owned).findResumableRun('conv_1', 'user_b')).toBeNull();
    const mine = await buildService(runDoc({ userId: 'user_a' })).findResumableRun('conv_1', 'user_a');
    expect(mine?.runId).toBe('run_1');
  });

  it('still resumes an ownerless run — the eval-bypass path predates auth wiring', async () => {
    const run = await buildService(runDoc()).findResumableRun('conv_1', 'user_a');
    expect(run?.runId).toBe('run_1');
  });

  it('queries only for genuinely unfinished statuses', async () => {
    const model = {
      findOne: jest.fn().mockReturnValue({ sort: jest.fn().mockResolvedValue(null) }),
    };
    await new AgentRunStateService(model as never).findResumableRun('conv_1');
    expect(model.findOne).toHaveBeenCalledWith({
      conversationId: 'conv_1',
      status: { $in: ['awaiting_decision', 'running'] },
    });
  });
});
