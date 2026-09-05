import { LlmRouterService } from '../src/excel-ai/services/llm-router.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { Tier2GenerateVerifyService } from '../src/excel-ai/services/tier2-generate-verify.service';
import { ExecutorAgent } from '../src/agents/executor.agent';
import { VerifierAgent } from '../src/agents/verifier.agent';
import { ToolBridgeService } from '../src/agents/tool-bridge.service';
import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import { AppConfigService } from '../src/config/app-config.service';
import { RouterInput } from '../src/excel-ai/types/router.types';
import { WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Plumbing for the gated model-swap eval: three tiers (Router, Tier 2
 * generate, and — unchanged — Tier 3 Planner/Executor) must be independently
 * overridable without any tier-routing/classification logic changing. These
 * tests pin the wiring, not the eval result itself (that's run manually via
 * eval/run-live-eval.ts against real models, not asserted in CI).
 */

describe('LlmRouterService — OPENROUTER_MODEL_ROUTER override', () => {
  const baseInput: RouterInput = {
    // Deliberately ambiguous — must not match any regex fast lane,
    // classifyComplexity, hasWriteIntent, or the data-query lane, so the
    // request actually reaches callLlmRouter() rather than short-circuiting.
    message: 'xyz123 do the thing with the stuff please',
    mode: 'action',
    sheetHeaders: ['A', 'B'],
    activeSheet: 'Sheet1',
  };

  function buildService(model: string | undefined, complete: jest.Mock): LlmRouterService {
    const openRouter = { complete } as unknown as OpenRouterService;
    const config = {
      openRouterModelRouter: model ?? 'openai/gpt-5-mini',
    } as unknown as AppConfigService;
    return new LlmRouterService(openRouter, config);
  }

  it('passes config.openRouterModelRouter as the explicit model on the LLM call', async () => {
    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({ route: 'write', complexity: 2, confidence: 0.8, reasoning: 'ok' }),
    );
    const service = buildService('inception/mercury-2.5-preview', complete);

    await service.route(baseInput);

    expect(complete).toHaveBeenCalledTimes(1);
    const call = complete.mock.calls[0][0] as { model?: string; tier?: string };
    expect(call.model).toBe('inception/mercury-2.5-preview');
    // tier is still passed as a fallback path inside complete() — model takes
    // priority there, so leaving it doesn't reintroduce the shared-LOW coupling.
    expect(call.tier).toBe('low');
  });

  it('is a no-op when OPENROUTER_MODEL_ROUTER is unset (falls back to openRouterModelLow)', async () => {
    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({ route: 'write', complexity: 2, confidence: 0.8, reasoning: 'ok' }),
    );
    const service = buildService('openai/gpt-5-mini', complete);

    await service.route(baseInput);

    const call = complete.mock.calls[0][0] as { model?: string };
    expect(call.model).toBe('openai/gpt-5-mini');
  });
});

describe('Tier2GenerateVerifyService — OPENROUTER_MODEL_TIER2_GENERATE override', () => {
  const workbookContext: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      {
        name: 'Sheet1',
        usedRange: 'A1:B2',
        rowCount: 2,
        columnCount: 2,
        values: [['Item', 'Qty'], ['Widget', 4]],
        formulas: [],
        numberFormats: [],
        structure: 'data_table',
        headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  function buildService(
    executorExecute: jest.Mock,
    tier2GenerateModel: string,
  ): Tier2GenerateVerifyService {
    const executor = { execute: executorExecute, retryStep: jest.fn() } as unknown as ExecutorAgent;
    const verifier = { verify: jest.fn() } as unknown as VerifierAgent;
    const formulaValidator = {
      checkNoHardcodedLiterals: jest.fn().mockReturnValue({ passed: true }),
    } as unknown as FormulaValidatorService;
    const toolBridge = {} as unknown as ToolBridgeService;
    const config = {
      openRouterModelTier2Generate: tier2GenerateModel,
    } as unknown as AppConfigService;
    return new Tier2GenerateVerifyService(executor, verifier, formulaValidator, toolBridge, config);
  }

  it('generateOnly() passes config.openRouterModelTier2Generate as ExecutorAgent.execute()\'s model override', async () => {
    const execute = jest.fn().mockResolvedValue({ subtaskId: 's1', actions: [], isDone: true });
    const service = buildService(execute, 'z-ai/glm-5.3-flash');

    await service.generateOnly('add a total column', 'FORMULA_GEN', workbookContext);

    expect(execute).toHaveBeenCalledTimes(1);
    // execute(subtask, context, previousActions, correlationId, usageTotals, modelOverride)
    const modelOverrideArg = execute.mock.calls[0][5];
    expect(modelOverrideArg).toBe('z-ai/glm-5.3-flash');
  });

  it('is a no-op when OPENROUTER_MODEL_TIER2_GENERATE is unset (falls back to openRouterModelHigh)', async () => {
    const execute = jest.fn().mockResolvedValue({ subtaskId: 's1', actions: [], isDone: true });
    const service = buildService(execute, 'openai/gpt-5');

    await service.generateOnly('add a total column', 'FORMULA_GEN', workbookContext);

    const modelOverrideArg = execute.mock.calls[0][5];
    expect(modelOverrideArg).toBe('openai/gpt-5');
  });
});

describe('ExecutorAgent — modelOverride param defaults to modelName (Tier 3 unaffected)', () => {
  it('execute() uses modelName when no override is passed (Tier 3 call shape)', async () => {
    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({ subtaskId: 's1', actions: [], isDone: true }),
    );
    const llm = { complete } as unknown as OpenRouterService;
    const config = { openRouterModelHigh: 'openai/gpt-5' } as unknown as AppConfigService;
    const agent = new ExecutorAgent(llm, config);

    await agent.execute(
      { id: 's1', description: 'noop', targetSheet: 'Sheet1', dependsOn: [], estimatedActions: 1 },
      { activeSheetName: 'Sheet1', sheets: [], namedRanges: [], tables: [] },
    );

    const call = complete.mock.calls[0][0] as { model?: string };
    expect(call.model).toBe('openai/gpt-5');
  });

  it('execute() uses the passed modelOverride when provided (Tier 2 call shape)', async () => {
    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({ subtaskId: 's1', actions: [], isDone: true }),
    );
    const llm = { complete } as unknown as OpenRouterService;
    const config = { openRouterModelHigh: 'openai/gpt-5' } as unknown as AppConfigService;
    const agent = new ExecutorAgent(llm, config);

    await agent.execute(
      { id: 's1', description: 'noop', targetSheet: 'Sheet1', dependsOn: [], estimatedActions: 1 },
      { activeSheetName: 'Sheet1', sheets: [], namedRanges: [], tables: [] },
      [],
      'corr_1',
      undefined,
      'z-ai/glm-5.3-flash',
    );

    const call = complete.mock.calls[0][0] as { model?: string };
    expect(call.model).toBe('z-ai/glm-5.3-flash');
  });
});
