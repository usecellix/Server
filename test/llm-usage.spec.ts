import { AppConfigService } from '../src/config/app-config.service';
import { ModelRouter } from '../src/excel-ai/llm/model-router';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { writeSseEvent } from '../src/excel-ai/utils/sse.util';
import { llmCallerFromStack } from '../src/llm-usage/llm-caller.util';
import {
  currentLlmUsageContext,
  runWithLlmUsageContext,
  updateLlmUsageContext,
} from '../src/llm-usage/llm-usage.context';
import { LlmUsageService, type RecordLlmCallInput } from '../src/llm-usage/llm-usage.service';

// The SDK is ESM-only; only its constructor is reached because the network send is stubbed.
jest.mock('@openrouter/sdk', () => ({ OpenRouter: class {} }));

const flush = () => new Promise((resolve) => setImmediate(resolve));

function buildUsageService() {
  const callModel = { create: jest.fn().mockResolvedValue({}) };
  const promptModel = {
    updateOne: jest.fn().mockReturnValue({ exec: jest.fn().mockResolvedValue({}) }),
  };
  const service = new LlmUsageService(promptModel as never, callModel as never);
  return { service, callModel, promptModel };
}

function baseCall(overrides: Partial<RecordLlmCallInput> = {}): RecordLlmCallInput {
  return {
    model: 'openai/gpt-5',
    caller: 'planner',
    attempt: 1,
    streaming: false,
    promptTokens: 1000,
    completionTokens: 200,
    reasoningTokens: 50,
    cachedTokens: 400,
    costUsd: 0.0123,
    costEstimated: false,
    latencyMs: 2500,
    success: true,
    ...overrides,
  };
}

describe('llmCallerFromStack', () => {
  it('skips OpenRouterService frames and names the calling agent', () => {
    const stack = [
      'Error',
      '    at OpenRouterService.complete (E:\\cellix\\cellix_backend\\src\\excel-ai\\services\\openrouter.service.ts:155:40)',
      '    at PlannerAgent.expandPhase (E:\\cellix\\cellix_backend\\src\\agents\\planner.agent.ts:812:31)',
    ].join('\n');
    expect(llmCallerFromStack(stack)).toBe('planner');
  });

  it('handles compiled .js paths and strips the .service suffix', () => {
    const stack = [
      'Error',
      '    at OpenRouterService.complete (/app/dist/excel-ai/services/openrouter.service.js:120:20)',
      '    at LlmRouterService.route (/app/dist/excel-ai/services/llm-router.service.js:88:9)',
    ].join('\n');
    expect(llmCallerFromStack(stack)).toBe('llm-router');
  });

  it('falls back to unknown when nothing in the stack is ours', () => {
    expect(llmCallerFromStack('Error\n    at node:internal/process/task_queues:95:5')).toBe('unknown');
    expect(llmCallerFromStack(undefined)).toBe('unknown');
  });
});

describe('LlmUsageService', () => {
  it('attributes a call to the prompt in context and increments its totals', async () => {
    const { service, callModel, promptModel } = buildUsageService();

    await runWithLlmUsageContext(
      { promptId: 'p1', userId: 'u1', conversationId: 'c1', route: 'write', tier: 3 },
      async () => {
        await Promise.resolve();
        service.recordCall(baseCall());
      },
    );
    await flush();

    expect(callModel.create).toHaveBeenCalledWith(
      expect.objectContaining({
        promptId: 'p1',
        userId: 'u1',
        conversationId: 'c1',
        caller: 'planner',
        totalTokens: 1200,
        costUsd: 0.0123,
      }),
    );
    const [filter, update, options] = promptModel.updateOne.mock.calls[0];
    expect(filter).toEqual({ promptId: 'p1' });
    expect(options).toEqual({ upsert: true });
    expect(update.$inc).toEqual(
      expect.objectContaining({
        llmCalls: 1,
        failedCalls: 0,
        promptTokens: 1000,
        completionTokens: 200,
        totalTokens: 1200,
        costUsd: 0.0123,
      }),
    );
    expect(update.$addToSet).toEqual({ models: 'openai/gpt-5' });
    expect(update.$set).toEqual(expect.objectContaining({ route: 'write', tier: 3 }));
  });

  it('counts a failed call as failed', async () => {
    const { service, promptModel } = buildUsageService();
    runWithLlmUsageContext({ promptId: 'p1' }, () =>
      service.recordCall(baseCall({ success: false, costUsd: 0, errorStatus: 429 })),
    );
    await flush();
    expect(promptModel.updateOne.mock.calls[0][1].$inc.failedCalls).toBe(1);
  });

  it('records calls made outside a request as unattributed, without touching any prompt', async () => {
    const { service, callModel, promptModel } = buildUsageService();
    service.recordCall(baseCall());
    await flush();
    expect(callModel.create).toHaveBeenCalledWith(expect.objectContaining({ promptId: null }));
    expect(promptModel.updateOne).not.toHaveBeenCalled();
  });

  it('does not create a prompt row while a /continue has not resolved its prompt id', async () => {
    const { service, promptModel } = buildUsageService();
    runWithLlmUsageContext({ promptId: '' }, () => {
      service.recordCall(baseCall());
      service.endRequest(currentLlmUsageContext()!, { durationMs: 10 });
    });
    await flush();
    expect(promptModel.updateOne).not.toHaveBeenCalled();
  });

  it('marks the request as an error when one was sent to the user over SSE', async () => {
    const { service, promptModel } = buildUsageService();
    const reply = { raw: { write: jest.fn() } } as never;
    const context = { promptId: 'p1' };
    runWithLlmUsageContext(context, () => {
      writeSseEvent(reply, 'error', { message: 'Planner failed to produce a plan' });
    });
    service.endRequest(context, { durationMs: 1234 });
    await flush();

    const update = promptModel.updateOne.mock.calls[0][1];
    expect(update.$set.lastOutcome).toBe('error');
    expect(update.$set.lastError).toBe('Planner failed to produce a plan');
    expect(update.$inc).toEqual({ requestCount: 1, requestDurationMs: 1234 });
  });

  it('retries once when a concurrent upsert races on the unique prompt id', async () => {
    const { service, promptModel } = buildUsageService();
    promptModel.updateOne
      .mockReturnValueOnce({ exec: jest.fn().mockRejectedValue({ code: 11000 }) })
      .mockReturnValue({ exec: jest.fn().mockResolvedValue({}) });
    runWithLlmUsageContext({ promptId: 'p1' }, () => service.recordCall(baseCall()));
    await flush();
    await flush();
    expect(promptModel.updateOne).toHaveBeenCalledTimes(2);
  });

  it('lets a request re-point its prompt id after it started', () => {
    runWithLlmUsageContext({ promptId: '' }, () => {
      updateLlmUsageContext({ promptId: 'original-trace', conversationId: 'c9' });
      expect(currentLlmUsageContext()).toEqual({ promptId: 'original-trace', conversationId: 'c9' });
    });
  });
});

describe('OpenRouterService usage recording', () => {
  function buildService() {
    const config = {
      openRouterApiKey: 'test-key',
      openRouterHttpReferer: 'http://localhost',
      openRouterModelLow: 'openai/gpt-5-mini',
      openRouterModelMedium: 'openai/gpt-5-mini',
      openRouterModelHigh: 'openai/gpt-5',
    } as unknown as AppConfigService;
    const modelRouter = { markRateLimited: jest.fn() } as unknown as ModelRouter;
    const usage = { recordCall: jest.fn() } as unknown as LlmUsageService & {
      recordCall: jest.Mock;
    };
    const service = new OpenRouterService(config, modelRouter, usage);
    const send = jest.spyOn(
      service as unknown as { sendChatCompletionRaw: (...args: unknown[]) => Promise<unknown> },
      'sendChatCompletionRaw',
    );
    return { service, usage, send };
  }

  it('records the provider-billed cost, served model and the calling file', async () => {
    const { service, usage, send } = buildService();
    send.mockResolvedValueOnce({
      id: 'gen-1',
      model: 'openai/gpt-5-2026-08',
      choices: [{ message: { content: '{"ok":true}' }, finishReason: 'stop' }],
      usage: {
        promptTokens: 900,
        completionTokens: 100,
        totalTokens: 1000,
        cost: 0.0042,
        completionTokensDetails: { reasoningTokens: 30 },
        promptTokensDetails: { cachedTokens: 512 },
      },
    });

    await service.complete({ systemPrompt: 's', userMessage: 'u', tier: 'high' });

    expect(usage.recordCall).toHaveBeenCalledTimes(1);
    expect(usage.recordCall).toHaveBeenCalledWith(
      expect.objectContaining({
        model: 'openai/gpt-5',
        servedModel: 'openai/gpt-5-2026-08',
        generationId: 'gen-1',
        caller: 'llm-usage.spec',
        attempt: 1,
        promptTokens: 900,
        completionTokens: 100,
        reasoningTokens: 30,
        cachedTokens: 512,
        costUsd: 0.0042,
        costEstimated: false,
        success: true,
        finishReason: 'stop',
      }),
    );
  });

  it('records every network attempt, numbering retries, including the failed one', async () => {
    const { service, usage, send } = buildService();
    const mandatory = Object.assign(
      new Error('Reasoning is mandatory for this endpoint and cannot be disabled.'),
      { status: 400 },
    );
    send.mockRejectedValueOnce(mandatory).mockResolvedValueOnce({
      choices: [{ message: { content: 'ok' }, finishReason: 'stop' }],
      usage: { promptTokens: 10, completionTokens: 5, totalTokens: 15, cost: 0.0001 },
    });

    await service.complete({
      systemPrompt: 's',
      userMessage: 'u',
      tier: 'low',
      reasoningEffort: 'none',
    });

    const calls = usage.recordCall.mock.calls.map(([c]) => c);
    expect(calls.map((c) => [c.attempt, c.success])).toEqual([
      [1, false],
      [2, true],
    ]);
    expect(calls[0]).toEqual(expect.objectContaining({ errorStatus: 400, costUsd: 0 }));
  });

  it('estimates cost from the model price table when the provider omits it', async () => {
    const { service, usage, send } = buildService();
    send.mockResolvedValueOnce({
      choices: [{ message: { content: 'ok' }, finishReason: 'stop' }],
      usage: { promptTokens: 1000, completionTokens: 1000, totalTokens: 2000 },
    });

    await service.complete({ systemPrompt: 's', userMessage: 'u', tier: 'high' });

    const [call] = usage.recordCall.mock.calls[0];
    expect(call.costEstimated).toBe(true);
    expect(call.costUsd).toBeGreaterThan(0);
  });
});
