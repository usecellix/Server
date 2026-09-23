import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { LlmRequestError } from '../src/excel-ai/errors/llm-request.error';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 3, acceptance criterion:
 *   "Simulated provider that rejects the Nth concurrent call: run completes
 *    with zero failed subtasks."
 *   "Iteration counter is unchanged by transient retries."
 *
 * The retry lives in `OpenRouterService.complete`, BELOW the agentic loop, so
 * the loop's iteration budget never sees it — that placement is the whole
 * point, and this test pins it.
 */
describe('OpenRouterService transient retry (Phase 3 acceptance, TASKS.md #286)', () => {
  function buildService(completeOnce: jest.Mock): OpenRouterService {
    const config = {
      openRouterApiKey: 'test-key',
      openRouterModelHigh: 'm',
      openRouterModelMedium: 'm',
      openRouterModelLow: 'm',
      openRouterHttpReferer: 'http://localhost',
    } as unknown as AppConfigService;
    const service = new OpenRouterService(config, {} as never);
    // Swap the single-shot call for the simulated provider; `complete` is the
    // wrapper under test and stays exactly as shipped.
    (service as unknown as { completeOnce: jest.Mock }).completeOnce = completeOnce;
    return service;
  }

  const CREDIT_CHECK_TIMEOUT = new LlmRequestError(
    402,
    'OpenRouter could not verify available credits for this request in time. Retry shortly.',
  );

  const opts = { systemPrompt: 's', userMessage: 'u' };

  it('recovers from the live credit-check timeout instead of failing the subtask', async () => {
    const completeOnce = jest
      .fn()
      .mockRejectedValueOnce(CREDIT_CHECK_TIMEOUT)
      .mockResolvedValue('{"ok":true}');

    const service = buildService(completeOnce);
    await expect(service.complete(opts as never)).resolves.toBe('{"ok":true}');
    expect(completeOnce).toHaveBeenCalledTimes(2);
  });

  it('rides out a provider that rejects the first TWO attempts', async () => {
    const completeOnce = jest
      .fn()
      .mockRejectedValueOnce(CREDIT_CHECK_TIMEOUT)
      .mockRejectedValueOnce(new LlmRequestError(429, 'rate limited'))
      .mockResolvedValue('{"ok":true}');

    const service = buildService(completeOnce);
    await expect(service.complete(opts as never)).resolves.toBe('{"ok":true}');
    expect(completeOnce).toHaveBeenCalledTimes(3);
  });

  it('gives up after a bounded number of attempts rather than retrying forever', async () => {
    const completeOnce = jest.fn().mockRejectedValue(CREDIT_CHECK_TIMEOUT);
    const service = buildService(completeOnce);

    await expect(service.complete(opts as never)).rejects.toThrow(/credits/i);
    expect(completeOnce).toHaveBeenCalledTimes(3);
  });

  it('does NOT retry a permanent fault — an empty wallet fails fast, as before', async () => {
    const completeOnce = jest
      .fn()
      .mockRejectedValue(new LlmRequestError(402, 'Insufficient credits. Add more to continue.'));
    const service = buildService(completeOnce);

    await expect(service.complete(opts as never)).rejects.toThrow(/insufficient/i);
    expect(completeOnce).toHaveBeenCalledTimes(1);
  });

  it('does NOT retry an auth failure', async () => {
    const completeOnce = jest.fn().mockRejectedValue(new LlmRequestError(401, 'bad key'));
    const service = buildService(completeOnce);

    await expect(service.complete(opts as never)).rejects.toThrow(/bad key/i);
    expect(completeOnce).toHaveBeenCalledTimes(1);
  });

  it('a successful first call costs no retries at all', async () => {
    const completeOnce = jest.fn().mockResolvedValue('{"ok":true}');
    const service = buildService(completeOnce);

    await expect(service.complete(opts as never)).resolves.toBe('{"ok":true}');
    expect(completeOnce).toHaveBeenCalledTimes(1);
  });
});
