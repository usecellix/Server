import { LlmRouterService } from '../src/excel-ai/services/llm-router.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';

describe('LlmRouterService.classifyIntent (CHITCHAT route)', () => {
  let openRouter: jest.Mocked<Pick<OpenRouterService, 'complete'>>;
  let config: Pick<AppConfigService, 'openRouterModelLow'>;
  let service: LlmRouterService;

  beforeEach(() => {
    openRouter = {
      complete: jest.fn(),
    };
    config = { openRouterModelLow: 'openai/gpt-5-nano' };
    service = new LlmRouterService(
      openRouter as unknown as OpenRouterService,
      config as AppConfigService,
    );
  });

  it.each(['hi', 'hello', 'thanks!', 'who are you', 'what can you do'])(
    'classifies "%s" as CHITCHAT without invoking tier-classification logic',
    async (message) => {
      openRouter.complete.mockResolvedValue('CHITCHAT');
      const routeSpy = jest.spyOn(service, 'route');

      const label = await service.classifyIntent(message);

      expect(label).toBe('CHITCHAT');
      expect(routeSpy).not.toHaveBeenCalled();
    },
  );

  it('requests text (not JSON) mode — the default json_object mode is what made this classifier dead in production (TASKS.md #228)', async () => {
    openRouter.complete.mockResolvedValue('CHITCHAT');

    await service.classifyIntent('hi');

    expect(openRouter.complete).toHaveBeenCalledWith(
      expect.objectContaining({ responseFormat: 'text' }),
    );
  });

  it.each([
    ['{"label":"TASK"}', 'TASK'],
    ['{"label": "CHITCHAT"}', 'CHITCHAT'],
    ['```\nTASK\n```', 'TASK'],
    ['Label: CHITCHAT.', 'CHITCHAT'],
  ] as const)(
    'still classifies correctly when the model wraps the label as %s (regression for the JSON-wrapped dead-classifier bug)',
    async (raw, expected) => {
      openRouter.complete.mockResolvedValue(raw);

      const label = await service.classifyIntent('some message');

      expect(label).toBe(expected);
    },
  );

  it('does NOT skip reasoning.effort for a model that requires it (regression for the guaranteed-failing effort:none attempt)', async () => {
    const glmConfig = { openRouterModelLow: 'z-ai/glm-5.3-flash' };
    const glmService = new LlmRouterService(
      openRouter as unknown as OpenRouterService,
      glmConfig as AppConfigService,
    );
    openRouter.complete.mockResolvedValue('CHITCHAT');

    await glmService.classifyIntent('hi');

    expect(openRouter.complete).toHaveBeenCalledWith(
      expect.objectContaining({ reasoningEffort: 'low' }),
    );
  });

  it('still sends effort:none for a model that has no reasoning-mandatory quirk', async () => {
    openRouter.complete.mockResolvedValue('CHITCHAT');

    await service.classifyIntent('hi');

    expect(openRouter.complete).toHaveBeenCalledWith(
      expect.objectContaining({ reasoningEffort: 'none' }),
    );
  });

  it.each(['add a column', 'sort by date'])(
    'classifies "%s" as TASK and leaves existing tier logic unaffected',
    async (message) => {
      openRouter.complete.mockResolvedValue('TASK');

      const label = await service.classifyIntent(message);
      expect(label).toBe('TASK');

      // Existing route() logic runs exactly as before — unaffected by classifyIntent.
      openRouter.complete.mockResolvedValue(
        JSON.stringify({ route: 'write', confidence: 0.9, reasoning: 'test', complexity: 1 }),
      );
      const decision = await service.route({
        message,
        mode: 'action',
        sheetHeaders: ['Date', 'Amount'],
        activeSheet: 'Sheet1',
      });
      expect(decision.route).toBe('write');
    },
  );

  it('falls back to TASK and logs a warning when the classifier call throws', async () => {
    openRouter.complete.mockRejectedValue(new Error('OpenRouter timeout'));
    const warnSpy = jest.spyOn(service['logger'], 'warn').mockImplementation(() => undefined);

    const label = await service.classifyIntent('hi');

    expect(label).toBe('TASK');
    expect(warnSpy).toHaveBeenCalled();
  });

  it('falls back to TASK when the classifier returns an unrecognized label', async () => {
    openRouter.complete.mockResolvedValue('MAYBE');
    const warnSpy = jest.spyOn(service['logger'], 'warn').mockImplementation(() => undefined);

    const label = await service.classifyIntent('some ambiguous message');

    expect(label).toBe('TASK');
    expect(warnSpy).toHaveBeenCalled();
  });
});
