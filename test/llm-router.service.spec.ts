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
