import { OrchestratorService } from '../src/agents/orchestrator.service';
import { PlannerAgent } from '../src/agents/planner.agent';
import { ExecutorAgent } from '../src/agents/executor.agent';
import { VerifierAgent } from '../src/agents/verifier.agent';
import { AgenticLoopService } from '../src/agents/agenticLoop.service';
import { SseEmitter } from '../src/agents/sse.emitter';
import { AuditService } from '../src/audit/audit.service';
import { OpenRouterService, LlmCallTelemetry } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { WorkbookContext, SubTask } from '../src/agents/types/agent.types';
import { createUsageAccumulator } from '../src/agents/utils/usage-accumulator.util';

/**
 * conversation.service.ts's `streamWithOrchestrator`/`streamWithPlanner`
 * `finally` blocks call `auditService.logLLMCall({ ..., promptTokens:
 * telemetry.usage?.promptTokens ?? 0, completionTokens: telemetry.usage?.
 * completionTokens ?? 0, ... })` — and `telemetry` was NEVER populated for
 * the Tier-3 path, so every audit_logs row for it reported 0/0 regardless of
 * what the Planner/Executor/Verifier actually used. This test exercises the
 * real fix end to end: a real `PlannerAgent.plan()` call (only its LLM client
 * mocked, with known usage numbers) through the real `OrchestratorService`,
 * populating a real `LlmCallTelemetry` out-param exactly as
 * conversation.service.ts does, then feeding those exact values into a real
 * `AuditService.logLLMCall()` (only its Mongoose model mocked) — and asserts
 * the resulting audit_logs write is non-zero.
 */

function sheet(name: string, values: unknown[][]): WorkbookContext['sheets'][number] {
  const rowCount = values.length;
  const columnCount = values[0]?.length ?? 0;
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

const context: WorkbookContext = {
  activeSheetName: 'Dashboard',
  sheets: [sheet('Dashboard', [['A', 'B'], [1, 2]])],
  namedRanges: [],
  tables: [],
};

const KNOWN_PLANNER_USAGE = {
  promptTokens: 4321,
  completionTokens: 987,
  totalTokens: 5308,
};

function buildRealPlannerAgent(): PlannerAgent {
  // Only the LLM client is mocked — this is a real PlannerAgent.plan() call,
  // real prompt-building, real parsing, real usage accumulation.
  const llm = {
    complete: jest
      .fn()
      .mockImplementation((opts: { outcome?: { truncated?: boolean; usage?: unknown } }) => {
        if (opts.outcome) {
          opts.outcome.truncated = false;
          opts.outcome.usage = KNOWN_PLANNER_USAGE;
        }
        return Promise.resolve(
          JSON.stringify({
            subtasks: [
              {
                id: 's1',
                description: 'Create chart on Dashboard',
                targetSheet: 'Dashboard',
                dependsOn: [],
                estimatedActions: 1,
              },
            ],
            clarificationsNeeded: [],
            confidence: 'high',
            reasoning: 'ok',
          }),
        );
      }),
  } as unknown as OpenRouterService;

  const config = {
    openRouterModelHigh: 'openai/gpt-5',
    openRouterModelPlanner: 'openai/gpt-5',
    openRouterModelLow: 'openai/gpt-5-mini',
    openRouterModelMedium: 'openai/gpt-5-mini',
  } as unknown as AppConfigService;

  return new PlannerAgent(llm, config);
}

/** Executor/Verifier are stubbed here — this test isolates the Planner's
 * contribution to the shared usage accumulator, per the specific ask. */
function buildStubAgenticLoop(): AgenticLoopService {
  return {
    run: jest.fn().mockResolvedValue({
      actions: [],
      iterationsRun: 1,
      verifierPassed: true,
      completedSubtasks: [],
      failedSubtask: null,
      partialProgress: false,
    }),
  } as unknown as AgenticLoopService;
}

function buildAuditServiceWithMockedModel(): { service: AuditService; create: jest.Mock } {
  const create = jest.fn().mockResolvedValue(undefined);
  // Constructor args: auditEntryModel, auditLogModel, changeSetService — only
  // auditLogModel is exercised by logLLMCall.
  const service = new AuditService({} as never, { create } as never, {} as never);
  return { service, create };
}

describe('Spec 16 fix #5 — Tier-3 audit_logs telemetry population', () => {
  it('a real PlannerAgent.plan() call with known usage results in a non-zero audit_logs entry', async () => {
    const planner = buildRealPlannerAgent();
    const agenticLoop = buildStubAgenticLoop();
    const orchestrator = new OrchestratorService(planner, agenticLoop);
    const { service: auditService, create } = buildAuditServiceWithMockedModel();

    const telemetry: LlmCallTelemetry = { provider: 'openrouter', modelTier: 'high' };
    const emitter = new SseEmitter(() => {});

    const result = await orchestrator.runDetailed(
      { prompt: 'Create a chart on Dashboard', context, conversationHistory: [], complexity: 3 },
      emitter,
      telemetry,
    );

    // Sanity: the run actually produced a plan (not a clarification bail).
    expect(result.clarificationRequested).toBe(false);

    // The out-param is populated with the REAL usage from the mocked LLM response.
    expect(telemetry.usage?.promptTokens).toBe(KNOWN_PLANNER_USAGE.promptTokens);
    expect(telemetry.usage?.completionTokens).toBe(KNOWN_PLANNER_USAGE.completionTokens);
    expect(telemetry.model).toBe('openai/gpt-5');

    // Exactly what conversation.service.ts's finally block does at both
    // streamWithOrchestrator (line ~2594) and streamWithPlanner (~2711).
    await auditService.logLLMCall({
      traceId: 'trace-1',
      model: telemetry.model ?? 'orchestrator',
      tier: (telemetry.modelTier ?? 'high') as 'low' | 'medium' | 'high',
      intent: 'write',
      promptTokens: telemetry.usage?.promptTokens ?? 0,
      completionTokens: telemetry.usage?.completionTokens ?? 0,
      latencyMs: 1234,
      success: true,
    });

    expect(create).toHaveBeenCalledTimes(1);
    const written = create.mock.calls[0][0] as {
      llmModel: string;
      promptTokens: number;
      completionTokens: number;
      totalTokens: number;
      estimatedCostUsd: number;
    };

    // The actual assertion this fix exists for: non-zero, not the old 0/0/0.
    expect(written.promptTokens).toBe(KNOWN_PLANNER_USAGE.promptTokens);
    expect(written.completionTokens).toBe(KNOWN_PLANNER_USAGE.completionTokens);
    expect(written.totalTokens).toBe(
      KNOWN_PLANNER_USAGE.promptTokens + KNOWN_PLANNER_USAGE.completionTokens,
    );
    expect(written.estimatedCostUsd).toBeGreaterThan(0);
    expect(written.llmModel).toBe('openai/gpt-5');
  });

  it('still reports 0 usage (not a crash) when telemetry is never wired by the caller', async () => {
    const planner = buildRealPlannerAgent();
    const agenticLoop = buildStubAgenticLoop();
    const orchestrator = new OrchestratorService(planner, agenticLoop);
    const emitter = new SseEmitter(() => {});

    // No telemetry out-param passed — must not throw, and behaves exactly
    // like before this fix (an untouched call site is unaffected).
    await expect(
      orchestrator.runDetailed(
        { prompt: 'Create a chart on Dashboard', context, conversationHistory: [], complexity: 3 },
        emitter,
      ),
    ).resolves.toMatchObject({ clarificationRequested: false });
  });
});

describe('Spec 16 fix #5 — Executor and Verifier also accumulate into the shared usageTotals', () => {
  const subtask: SubTask = {
    id: 's1',
    description: 'Create chart on Dashboard',
    targetSheet: 'Dashboard',
    dependsOn: [],
    estimatedActions: 1,
  };

  it('ExecutorAgent.execute() adds its real usage to the accumulator', async () => {
    const usage = { promptTokens: 1500, completionTokens: 300, totalTokens: 1800 };
    const complete = jest
      .fn()
      .mockImplementation((opts: { outcome?: { usage?: unknown } }) => {
        if (opts.outcome) opts.outcome.usage = usage;
        return Promise.resolve(JSON.stringify({ subtaskId: 's1', actions: [], isDone: true }));
      });
    const config = { openRouterModelHigh: 'openai/gpt-5' } as unknown as AppConfigService;
    const agent = new ExecutorAgent({ complete } as unknown as OpenRouterService, config);

    const totals = createUsageAccumulator();
    await agent.execute(subtask, context, [], 'corr_exec', totals);

    expect(totals.promptTokens).toBe(1500);
    expect(totals.completionTokens).toBe(300);
  });

  it('VerifierAgent.verify() adds its real usage to the SAME accumulator as a prior Executor call', async () => {
    const executorUsage = { promptTokens: 1500, completionTokens: 300, totalTokens: 1800 };
    const verifierUsage = { promptTokens: 900, completionTokens: 150, totalTokens: 1050 };

    const executorComplete = jest
      .fn()
      .mockImplementation((opts: { outcome?: { usage?: unknown } }) => {
        if (opts.outcome) opts.outcome.usage = executorUsage;
        return Promise.resolve(JSON.stringify({ subtaskId: 's1', actions: [], isDone: true }));
      });
    const verifierComplete = jest
      .fn()
      .mockImplementation((opts: { outcome?: { usage?: unknown } }) => {
        if (opts.outcome) opts.outcome.usage = verifierUsage;
        return Promise.resolve(
          JSON.stringify({
            passed: true,
            feedback: 'ok',
            issues: [],
            subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'ok', issues: [] }],
          }),
        );
      });

    const execConfig = { openRouterModelHigh: 'openai/gpt-5' } as unknown as AppConfigService;
    const verifyConfig = { openRouterModelMedium: 'openai/gpt-5-mini' } as unknown as AppConfigService;
    const executor = new ExecutorAgent({ complete: executorComplete } as unknown as OpenRouterService, execConfig);
    const verifier = new VerifierAgent({ complete: verifierComplete } as unknown as OpenRouterService, verifyConfig);

    // One shared accumulator, exactly as OrchestratorService/AgenticLoopService
    // thread the SAME object through Planner, every Executor call, and every
    // Verifier call for one run.
    const totals = createUsageAccumulator();
    await executor.execute(subtask, context, [], 'corr_exec', totals);
    await verifier.verify('Create a chart on Dashboard', [subtask], { s1: [] }, context, undefined, 'corr_verify', totals);

    expect(totals.promptTokens).toBe(executorUsage.promptTokens + verifierUsage.promptTokens);
    expect(totals.completionTokens).toBe(executorUsage.completionTokens + verifierUsage.completionTokens);
  });
});
