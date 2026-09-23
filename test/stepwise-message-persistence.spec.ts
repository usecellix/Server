import { ConversationService } from '../src/excel-ai/services/conversation.service';
import { Tier0DirectService } from '../src/excel-ai/services/tier0-direct.service';
import { FastifyReply } from 'fastify';
import * as sseUtil from '../src/excel-ai/utils/sse.util';

/**
 * TASKS.md #267 — the stepwise (multi-wave) Tier 3 execution path never
 * called `saveMessage` for ANY wave, in ANY run. Every wave's 'actions' SSE
 * event rendered live in the browser, but nothing was ever written to the
 * conversation's persisted `messages` array — so `GET /conversation/:id`
 * (used by `openConversationFromHistory` on the client) returned only the
 * original user message. Reopening a completed multi-sheet build from
 * history showed a tab with the right title and a completely empty body:
 * `messagesToTurns` had a `turn.userMessage` but zero blocks to render.
 *
 * Same direct-private-method-call harness as
 * credit-tier3-stepwise-debit.spec.ts / mode-plan-only.spec.ts.
 */
describe('ConversationService — stepwise message persistence (TASKS.md #267)', () => {
  let service: ConversationService;
  let conversationModel: { updateOne: jest.Mock };
  let agentRunState: {
    markStatus: jest.Mock;
    summarizeSkipped: jest.Mock;
    nextExecutableWave: jest.Mock;
    recordWaveResult: jest.Mock;
  };
  let orchestrator: { runStepwiseWave: jest.Mock };
  let changeSetService: { createPreview: jest.Mock };
  let creditLedger: { debit: jest.Mock };
  const emittedEvents: Array<{ event: string; data: Record<string, unknown> }> = [];

  const reply = {} as FastifyReply;

  beforeEach(() => {
    emittedEvents.length = 0;
    conversationModel = { updateOne: jest.fn().mockResolvedValue(undefined) };
    agentRunState = {
      markStatus: jest.fn().mockResolvedValue(undefined),
      summarizeSkipped: jest.fn().mockReturnValue([]),
      nextExecutableWave: jest.fn(),
      recordWaveResult: jest.fn().mockResolvedValue(undefined),
    };
    orchestrator = { runStepwiseWave: jest.fn() };
    changeSetService = { createPreview: jest.fn() };
    creditLedger = { debit: jest.fn().mockResolvedValue({ debited: false }) };

    service = new ConversationService(
      conversationModel as never, // 1 conversationModel
      {} as never, // 2 sheetAnalyzer
      {} as never, // 3 engine
      {} as never, // 4 auditService
      changeSetService as never, // 5 changeSetService
      {} as never, // 6 openRouter
      orchestrator as never, // 7 orchestrator
      {} as never, // 8 llmRouter
      {} as never, // 9 chitchat
      {} as never, // 10 contextCache
      {} as never, // 11 dataQuery
      {} as never, // 12 findExport
      {} as never, // 13 formulaAnalyzer
      {} as never, // 14 toolBridge
      {} as never, // 15 smartDataQuery
      new Tier0DirectService(), // 16 tier0Direct
      {} as never, // 17 tier1SingleAction
      {} as never, // 18 tier2GenerateVerify
      {} as never, // 19 structuredLogger
      { appendNode: jest.fn(), setMeta: jest.fn() } as never, // 20 workflowTrace
      {} as never, // 21 creditGate
      creditLedger as never, // 22 creditLedger
      agentRunState as never, // 23 agentRunState
    );

    jest.spyOn(service as never, 'markCompleted' as never).mockResolvedValue(undefined as never);
    jest.spyOn(sseUtil, 'endSseResponse').mockImplementation(() => undefined);
    // executeStepwiseWave calls createRequestAbortSignal(reply) — reply.raw
    // does not exist on the bare {} used here, so stub the real helper the
    // same way endSseResponse is stubbed above.
    jest.spyOn(sseUtil, 'createRequestAbortSignal').mockReturnValue({
      aborted: false,
    } as AbortSignal);
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  const emit = (event: string, data: Record<string, unknown>) => {
    emittedEvents.push({ event, data });
  };

  function callExecuteStepwiseWave(run: unknown) {
    return (service as any).executeStepwiseWave(run, reply, emit, {} as never, {} as never);
  }

  function callFinishStepwiseRun(run: unknown) {
    return (service as any).finishStepwiseRun(run, reply, emit);
  }

  function buildRun(overrides: Partial<Record<string, unknown>> = {}) {
    return {
      runId: 'run_1',
      conversationId: 'conv-1',
      prompt: 'build the ledger',
      context: { activeSheetName: 'Sheet1' },
      subtasks: [{ id: 's1', targetSheet: 'January', description: 'Create January', dependsOn: [] }],
      subtaskStates: [],
      waveIndex: -1,
      waveTotal: 1,
      changeSetIds: [],
      save: jest.fn().mockResolvedValue(undefined),
      ...overrides,
    };
  }

  it('persists a message carrying this wave\'s actions and changeSetId, matching the live "actions" SSE event', async () => {
    const run = buildRun();
    agentRunState.nextExecutableWave.mockReturnValue({
      waveIndex: 0,
      subtasks: run.subtasks,
    });
    orchestrator.runStepwiseWave.mockResolvedValue({
      actions: [{ type: 'ADD_SHEET', sheetName: 'January', name: 'January' }],
      completedSubtasks: [{ subtaskId: 's1', actions: [], verified: true }],
      failedSubtask: null,
      failedSubtasks: [],
      verifierPassed: true,
    });
    changeSetService.createPreview.mockResolvedValue({
      changeSetId: 'cs-1',
      changes: [],
      irreversibleActionTypes: [],
    });

    await callExecuteStepwiseWave(run);

    // The live card:
    const actionsEvent = emittedEvents.find((e) => e.event === 'actions');
    expect(actionsEvent?.data.changeSetId).toBe('cs-1');

    // The persisted twin that history reload rebuilds the card from:
    expect(conversationModel.updateOne).toHaveBeenCalledTimes(1);
    const [filter, update] = conversationModel.updateOne.mock.calls[0];
    expect(filter).toEqual({ conversationId: 'conv-1' });
    const pushedMessage = update.$push.messages;
    expect(pushedMessage.role).toBe('assistant');
    expect(pushedMessage.metadata.changeSetId).toBe('cs-1');
    expect(pushedMessage.metadata.actions).toEqual([
      { type: 'ADD_SHEET', sheetName: 'January', name: 'January' },
    ]);
  });

  it('persists the closing summary when the run finishes, so the reloaded thread does not end on a bare card', async () => {
    const run = buildRun({
      waveIndex: 2,
      waveTotal: 3,
      // A clean, consistent run: the plan targeted January and the run
      // actually created it. Without this the fixture describes a build whose
      // planned sheet was never created, and Phase 4 (TASKS.md #287) rightly
      // appends a gap note to the summary.
      subtaskStates: [
        {
          subtaskId: 's1',
          completed: true,
          actions: [{ type: 'ADD_SHEET', sheetName: 'January', name: 'January' }],
        },
      ],
    });
    await callFinishStepwiseRun(run);

    expect(conversationModel.updateOne).toHaveBeenCalledTimes(1);
    const pushedMessage = conversationModel.updateOne.mock.calls[0][1].$push.messages;
    expect(pushedMessage.role).toBe('assistant');
    expect(pushedMessage.type).toBe('answer');
    expect(pushedMessage.content).toBe('All steps applied.');
  });

  /**
   * TASKS.md #269 — `applyConsolidationPass` (#142) only ever ran inside
   * `finalizeActions`, which the stepwise path never calls. So a large build
   * (the only kind that goes stepwise) shipped a consolidated header on Main
   * that never filled in: "when i write the guest name and date ... it wont
   * reflect in the main sheet".
   *
   * It cannot see the pattern from one wave alone — the month CREATEs and
   * their headers live in earlier waves than Main's consolidated header — so
   * the pass must run over the accumulated run, not just this wave.
   */
  it('appends the consolidation formula using PRIOR waves as context, so Main reflects the month sheets', async () => {
    const schema = ['Unit No', 'Guest Name', 'Check In', 'Total Amount'];
    // Earlier waves: two month sheets, each created with the same schema.
    const priorActions = ['January', 'February'].map((month) => ({
      subtask: { id: `s_${month}`, targetSheet: month, description: `Create ${month}`, dependsOn: [] },
      actions: [
        { type: 'CREATE_SHEET', sheetName: month, name: month },
        {
          type: 'BATCH_SET',
          sheetName: month,
          operations: schema.map((label, i) => ({
            address: `${String.fromCharCode(65 + i)}1`,
            value: label,
          })),
        },
      ],
    }));

    const run = buildRun({
      excelCapabilities: { dynamicArrays: true },
      subtaskStates: priorActions.map((p) => ({
        subtaskId: p.subtask.id,
        actions: p.actions,
        completed: true,
      })),
      subtasks: [
        ...priorActions.map((p) => p.subtask),
        { id: 's_main', targetSheet: 'Main', description: 'Consolidated header', dependsOn: [] },
      ],
    });

    agentRunState.nextExecutableWave.mockReturnValue({
      waveIndex: 2,
      subtasks: [run.subtasks[run.subtasks.length - 1]],
    });
    // This wave writes Main's consolidated header: an origin column ("Month")
    // followed by exactly the month sheets' schema.
    orchestrator.runStepwiseWave.mockResolvedValue({
      actions: [
        { type: 'CREATE_SHEET', sheetName: 'Main', name: 'Main' },
        {
          type: 'BATCH_SET',
          sheetName: 'Main',
          operations: ['Month', ...schema].map((label, i) => ({
            address: `${String.fromCharCode(65 + i)}18`,
            value: label,
          })),
        },
      ],
      completedSubtasks: [{ subtaskId: 's_main', actions: [], verified: true }],
      failedSubtask: null,
      failedSubtasks: [],
      verifierPassed: true,
    });
    changeSetService.createPreview.mockResolvedValue({
      changeSetId: 'cs-consolidation',
      changes: [],
      irreversibleActionTypes: [],
    });

    await callExecuteStepwiseWave(run);

    const previewed = changeSetService.createPreview.mock.calls[0][0].actions as Array<{
      type: string;
      sheetName?: string;
      formula?: string;
    }>;
    const consolidation = previewed.find((a) => a.type === 'SET_FORMULA');
    expect(consolidation).toBeDefined();
    expect(consolidation!.sheetName).toBe('Main');
    // A live spilling formula that reads the month sheets, not a static copy.
    expect(consolidation!.formula).toContain('January');
    expect(consolidation!.formula).toContain('February');
  });

  it('a full 2-wave run persists 3 messages total across its /continue calls: one per wave, plus the closing summary', async () => {
    // Each wave is its own HTTP request in the real flow (the client calls
    // /continue after accepting) — executeStepwiseWave does not self-recurse
    // through successful waves, only through the empty-wave skip branch. So
    // this drives it the same way: one call per wave, reusing the same `run`
    // object a real request would reload between calls.
    const run = buildRun({
      waveTotal: 2,
      // Consistent with the plan: January was targeted AND created, so Phase 4's
      // reconciliation (TASKS.md #287) has nothing to report and the closing
      // summary stays exactly "All steps applied."
      subtaskStates: [
        {
          subtaskId: 's1',
          completed: true,
          actions: [{ type: 'ADD_SHEET', sheetName: 'January', name: 'January' }],
        },
      ],
    });

    orchestrator.runStepwiseWave.mockResolvedValue({
      actions: [{ type: 'BATCH_SET', sheetName: 'January', operations: [] }],
      completedSubtasks: [{ subtaskId: 's1', actions: [], verified: true }],
      failedSubtask: null,
      failedSubtasks: [],
      verifierPassed: true,
    });
    changeSetService.createPreview.mockResolvedValue({
      changeSetId: 'cs-x',
      changes: [],
      irreversibleActionTypes: [],
    });

    agentRunState.nextExecutableWave.mockReturnValueOnce({ waveIndex: 0, subtasks: run.subtasks });
    await callExecuteStepwiseWave(run); // wave 1 of 2

    agentRunState.nextExecutableWave.mockReturnValueOnce({ waveIndex: 1, subtasks: run.subtasks });
    await callExecuteStepwiseWave(run); // wave 2 of 2

    agentRunState.nextExecutableWave.mockReturnValueOnce(null);
    await callExecuteStepwiseWave(run); // nextExecutableWave returns null -> finishStepwiseRun

    expect(conversationModel.updateOne).toHaveBeenCalledTimes(3);
    const roles = conversationModel.updateOne.mock.calls.map((c) => c[1].$push.messages.role);
    expect(roles).toEqual(['assistant', 'assistant', 'assistant']);
    const lastMessage = conversationModel.updateOne.mock.calls[2][1].$push.messages;
    expect(lastMessage.content).toBe('All steps applied.');
    expect(lastMessage.metadata).toBeUndefined();
  });
});
