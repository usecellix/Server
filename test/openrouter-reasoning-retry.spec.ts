import { isReasoningMandatoryError } from '../src/excel-ai/utils/reasoning-mandatory.util';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { ModelRouter } from '../src/excel-ai/llm/model-router';

describe('isReasoningMandatoryError', () => {
  it('matches OpenRouter 400 reasoning-mandatory message', () => {
    const err = Object.assign(
      new Error('Reasoning is mandatory for this endpoint and cannot be disabled.'),
      { status: 400 },
    );
    expect(isReasoningMandatoryError(err, 400)).toBe(true);
  });

  it('matches cannot be disabled phrasing without explicit status', () => {
    expect(
      isReasoningMandatoryError(
        new Error('reasoning.effort cannot be disabled for this model'),
      ),
    ).toBe(true);
  });

  it('rejects other 400s', () => {
    expect(isReasoningMandatoryError(new Error('invalid request'), 400)).toBe(false);
  });

  it('rejects non-400 status even with matching message', () => {
    expect(
      isReasoningMandatoryError(
        new Error('Reasoning is mandatory for this endpoint and cannot be disabled.'),
        502,
      ),
    ).toBe(false);
  });
});

describe('OpenRouterService reasoning retry', () => {
  function buildService(): OpenRouterService {
    const config = {
      openRouterApiKey: 'test-key',
      openRouterHttpReferer: 'http://localhost',
      openRouterModelLow: 'openai/gpt-5-mini',
      openRouterModelMedium: 'openai/gpt-5-mini',
      openRouterModelHigh: 'openai/gpt-5',
    } as unknown as AppConfigService;
    const modelRouter = { markRateLimited: jest.fn() } as unknown as ModelRouter;
    return new OpenRouterService(config, modelRouter);
  }

  it('retries with effort=low when none is rejected as mandatory', async () => {
    const service = buildService();
    const mandatory = Object.assign(
      new Error('Reasoning is mandatory for this endpoint and cannot be disabled.'),
      { status: 400 },
    );

    const sendOnce = jest
      .spyOn(
        service as unknown as {
          sendChatCompletionOnce: (...args: unknown[]) => Promise<unknown>;
        },
        'sendChatCompletionOnce',
      )
      .mockRejectedValueOnce(mandatory)
      .mockResolvedValueOnce({
        choices: [{ message: { content: 'Total CGST is 100' }, finishReason: 'stop' }],
      });

    const result = await (
      service as unknown as {
        requestChatCompletion: (
          client: unknown,
          opts: Record<string, unknown>,
        ) => Promise<{ choices?: Array<{ message?: { content?: string } }> }>;
      }
    ).requestChatCompletion(
      {},
      {
        model: 'openai/gpt-5-mini',
        messages: [{ role: 'user', content: 'total CGST?' }],
        temperature: 0.2,
        maxCompletionTokens: 512,
        reasoningEffort: 'none',
        responseFormat: 'text',
      },
    );

    expect(result.choices?.[0]?.message?.content).toBe('Total CGST is 100');
    expect(sendOnce).toHaveBeenCalledTimes(2);
    expect(sendOnce.mock.calls[0][1]).toEqual(
      expect.objectContaining({ reasoningEffort: 'none' }),
    );
    expect(sendOnce.mock.calls[1][1]).toEqual(
      expect.objectContaining({ reasoningEffort: 'low' }),
    );
  });

  it('does not retry when effort is already low', async () => {
    const service = buildService();
    const mandatory = Object.assign(
      new Error('Reasoning is mandatory for this endpoint and cannot be disabled.'),
      { status: 400 },
    );

    jest
      .spyOn(
        service as unknown as {
          sendChatCompletionOnce: (...args: unknown[]) => Promise<unknown>;
        },
        'sendChatCompletionOnce',
      )
      .mockRejectedValueOnce(mandatory);

    await expect(
      (
        service as unknown as {
          requestChatCompletion: (
            client: unknown,
            opts: Record<string, unknown>,
          ) => Promise<unknown>;
        }
      ).requestChatCompletion(
        {},
        {
          model: 'openai/gpt-5-mini',
          messages: [],
          temperature: 0,
          maxCompletionTokens: 100,
          reasoningEffort: 'low',
          responseFormat: 'text',
        },
      ),
    ).rejects.toThrow(/Reasoning is mandatory/);
  });
});
