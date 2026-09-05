import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { ModelRouter } from '../src/excel-ai/llm/model-router';

/**
 * Regression test for the live-eval hang: a Tier 3 request sat server-side
 * indefinitely with no completion and no error, while a manual retry of the
 * identical request got a clean, fast 402 from OpenRouter. Root cause — no
 * call site set a request timeout, so a provider response that never resolves
 * (rejected in a way that doesn't surface, overloaded, or genuinely hung)
 * stalled forever. Fixed by passing the SDK's own `timeoutMs` option and
 * treating its RequestTimeoutError/RequestAbortedError as a transient network
 * error (retry-once / fallback), the same as a dropped connection.
 */
describe('OpenRouterService — request timeout wiring', () => {
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

  it('passes a bounded timeoutMs to client.chat.send for a non-streaming call', async () => {
    const service = buildService();
    const send = jest.fn().mockResolvedValue({
      choices: [{ message: { content: 'ok' }, finishReason: 'stop' }],
    });
    const fakeClient = { chat: { send } };

    await (
      service as unknown as {
        sendChatCompletionOnce: (client: unknown, opts: Record<string, unknown>) => Promise<unknown>;
      }
    ).sendChatCompletionOnce(fakeClient, {
      model: 'openai/gpt-5-mini',
      messages: [{ role: 'user', content: 'hi' }],
      temperature: 0.2,
      maxCompletionTokens: 512,
      reasoningEffort: 'low',
      responseFormat: 'text',
    });

    expect(send).toHaveBeenCalledTimes(1);
    const [, options] = send.mock.calls[0];
    expect(options).toBeDefined();
    expect(typeof options.timeoutMs).toBe('number');
    expect(options.timeoutMs).toBeGreaterThan(0);
  });

  it('recognizes RequestTimeoutError as a transient, retryable network error', () => {
    const service = buildService();
    const timeoutError = Object.assign(new Error('The request timed out'), {
      name: 'RequestTimeoutError',
    });

    const isTransient = (
      service as unknown as { isTransientNetworkError: (e: unknown) => boolean }
    ).isTransientNetworkError(timeoutError);

    expect(isTransient).toBe(true);
  });

  it('recognizes RequestAbortedError as a transient, retryable network error', () => {
    const service = buildService();
    const abortedError = Object.assign(new Error('The request was aborted'), {
      name: 'RequestAbortedError',
    });

    const isTransient = (
      service as unknown as { isTransientNetworkError: (e: unknown) => boolean }
    ).isTransientNetworkError(abortedError);

    expect(isTransient).toBe(true);
  });

  it('a timed-out non-streaming call is retried once, not left to hang', async () => {
    const service = buildService();
    const timeoutError = Object.assign(new Error('The request timed out'), {
      name: 'RequestTimeoutError',
    });
    const sendOnce = jest
      .spyOn(
        service as unknown as {
          sendChatCompletionOnce: (...args: unknown[]) => Promise<unknown>;
        },
        'sendChatCompletionOnce',
      )
      .mockRejectedValueOnce(timeoutError)
      .mockResolvedValueOnce({
        choices: [{ message: { content: 'recovered' }, finishReason: 'stop' }],
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
        messages: [{ role: 'user', content: 'hi' }],
        temperature: 0.2,
        maxCompletionTokens: 512,
        reasoningEffort: 'low',
        responseFormat: 'text',
      },
    );

    expect(sendOnce).toHaveBeenCalledTimes(2);
    expect(result.choices?.[0]?.message?.content).toBe('recovered');
  });

  it('does not silently swallow a non-transient error into a retry loop', async () => {
    const service = buildService();
    const authError = Object.assign(new Error('Invalid API key'), { status: 401 });
    const sendOnce = jest
      .spyOn(
        service as unknown as {
          sendChatCompletionOnce: (...args: unknown[]) => Promise<unknown>;
        },
        'sendChatCompletionOnce',
      )
      .mockRejectedValueOnce(authError);

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
          messages: [{ role: 'user', content: 'hi' }],
          temperature: 0.2,
          maxCompletionTokens: 512,
          reasoningEffort: 'low',
          responseFormat: 'text',
        },
      ),
    ).rejects.toThrow('Invalid API key');
    expect(sendOnce).toHaveBeenCalledTimes(1);
  });
});
