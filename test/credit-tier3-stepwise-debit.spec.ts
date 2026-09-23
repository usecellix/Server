import { ConversationService } from '../src/excel-ai/services/conversation.service';
import { Tier0DirectService } from '../src/excel-ai/services/tier0-direct.service';
import { FastifyReply } from 'fastify';
import * as sseUtil from '../src/excel-ai/utils/sse.util';

/**
 * Wiring test for the credit gate/debit hooks added around Tier 3's
 * STEPWISE execution path (CREDIT_SYSTEM.md CD-3/CD-4, TIER3_AGENTIC_BUILD —
 * credit-system-v2 session). A stepwise run spans several `/continue`
 * round-trips for what the user experiences as ONE request, so — unlike
 * Tier 2's single-dispatch debit (credit-write-route-integration.spec.ts) —
 * the debit must fire EXACTLY ONCE for the whole run, in `finishStepwiseRun`,
 * priced against the REAL total of subtasks that actually delivered actions
 * across every wave, not the plan's own (possibly larger) subtask count.
 *
 * Uses the same direct-private-method-call pattern as
 * credit-write-route-integration.spec.ts / mode-plan-only.spec.ts.
 */
describe('ConversationService — credit debit around stepwise Tier 3 completion', () => {
  let service: ConversationService;
  let creditLedger: { debit: jest.Mock };
  let agentRunState: { markStatus: jest.Mock; summarizeSkipped: jest.Mock };
  // TASKS.md #267 — finishStepwiseRun now persists a closing message, so
  // conversationModel must be a real-enough stub or that call throws.
  let conversationModel: { updateOne: jest.Mock };
  const emittedEvents: Array<{ event: string; data: Record<string, unknown> }> = [];

  const reply = {} as FastifyReply;

  function buildRun(overrides: Partial<{ userId?: string; subtaskStates: Array<{ completed: boolean }> }> = {}) {
    return {
      runId: 'run_1',
      conversationId: 'conv-1',
      userId: overrides.userId,
      waveIndex: 2,
      waveTotal: 3,
      subtaskStates:
        overrides.subtaskStates ??
        [{ completed: true }, { completed: true }, { completed: true }, { completed: false }],
    } as never;
  }

  beforeEach(() => {
    emittedEvents.length = 0;
    creditLedger = { debit: jest.fn() };
    agentRunState = {
      markStatus: jest.fn().mockResolvedValue(undefined),
      summarizeSkipped: jest.fn().mockReturnValue([]),
    };
    conversationModel = { updateOne: jest.fn().mockResolvedValue(undefined) };

    service = new ConversationService(
      conversationModel as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      new Tier0DirectService(),
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      creditLedger as never,
      agentRunState as never,
    );

    jest.spyOn(service as never, 'markCompleted' as never).mockResolvedValue(undefined as never);
    jest.spyOn(sseUtil, 'endSseResponse').mockImplementation(() => undefined);
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  const emit = (event: string, data: Record<string, unknown>) => {
    emittedEvents.push({ event, data });
  };

  function callFinishStepwiseRun(run: unknown) {
    return (service as any).finishStepwiseRun(run, reply, emit);
  }

  it('debits TIER3_AGENTIC_BUILD once, priced at the REAL count of subtasks that delivered actions across every wave', async () => {
    creditLedger.debit.mockResolvedValue({
      debited: true,
      balances: { planCredits: 2700, purchasedCredits: 0, oneTimeCredits: 0 },
    });

    const run = buildRun({
      userId: 'user-1',
      // 3 completed, 1 not (e.g. skipped/rejected) — the debit must count only
      // the 3 that actually delivered, not the plan's total of 4.
      subtaskStates: [{ completed: true }, { completed: true }, { completed: true }, { completed: false }],
    });

    await callFinishStepwiseRun(run);

    expect(creditLedger.debit).toHaveBeenCalledTimes(1);
    expect(creditLedger.debit).toHaveBeenCalledWith('user-1', 'TIER3_AGENTIC_BUILD', 3, {
      conversationId: 'conv-1',
    });

    const creditsEvent = emittedEvents.find((e) => e.event === 'credits');
    expect(creditsEvent?.data).toEqual({
      planCredits: 2700,
      purchasedCredits: 0,
      oneTimeCredits: 0,
      debited: 300,
      actionType: 'TIER3_AGENTIC_BUILD',
    });
  });

  it('skips the debit entirely when the run has no userId (eval-bypass path)', async () => {
    const run = buildRun({ userId: undefined });

    await callFinishStepwiseRun(run);

    expect(creditLedger.debit).not.toHaveBeenCalled();
    expect(emittedEvents.find((e) => e.event === 'credits')).toBeUndefined();
  });

  it('does not fail run completion when the debit loses a balance race (CD-4/CD-6)', async () => {
    creditLedger.debit.mockResolvedValue({ debited: false });

    const run = buildRun({ userId: 'user-1' });
    await callFinishStepwiseRun(run);

    expect(emittedEvents.find((e) => e.event === 'credits')).toBeUndefined();
    // The run still reports completion normally — a lost race is not surfaced as a failure.
    const endEvent = emittedEvents.find((e) => e.event === 'conversation_end');
    expect(endEvent).toBeDefined();
  });

  it('still reports the run as complete (conversation_end + wave_ready hasMore:false) alongside the debit', async () => {
    creditLedger.debit.mockResolvedValue({
      debited: true,
      balances: { planCredits: 2700, purchasedCredits: 0, oneTimeCredits: 0 },
    });

    const run = buildRun({ userId: 'user-1' });
    await callFinishStepwiseRun(run);

    const waveReadyEvent = emittedEvents.find((e) => e.event === 'wave_ready');
    expect(waveReadyEvent?.data).toEqual(
      expect.objectContaining({ runId: 'run_1', hasMore: false }),
    );
  });

  it('debits ZERO when every subtask across the run failed to deliver (an honest zero-cost outcome, not skipped entirely)', async () => {
    creditLedger.debit.mockResolvedValue({
      debited: true,
      balances: { planCredits: 3000, purchasedCredits: 0, oneTimeCredits: 0 },
    });

    const run = buildRun({
      userId: 'user-1',
      subtaskStates: [{ completed: false }, { completed: false }],
    });

    await callFinishStepwiseRun(run);

    expect(creditLedger.debit).toHaveBeenCalledWith('user-1', 'TIER3_AGENTIC_BUILD', 0, {
      conversationId: 'conv-1',
    });
  });
});
