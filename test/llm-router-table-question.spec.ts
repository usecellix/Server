import { LlmRouterService } from '../src/excel-ai/services/llm-router.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';

/**
 * TASKS.md #383 — a question about the rows, typed while the mode selector is
 * on Action, must reach the data lane (where it is computed in code) without
 * consulting the model router, which could hand it to the write pipeline and
 * end in "could not parse the AI response".
 */
describe('LlmRouterService — a table question typed in Action mode', () => {
  let openRouter: jest.Mocked<Pick<OpenRouterService, 'complete'>>;
  let service: LlmRouterService;

  beforeEach(() => {
    openRouter = { complete: jest.fn() };
    service = new LlmRouterService(
      openRouter as unknown as OpenRouterService,
      { openRouterModelLow: 'test/model' } as AppConfigService,
    );
  });

  const route = (message: string, mode: 'action' | 'ask') =>
    service.route({ message, mode } as Parameters<LlmRouterService['route']>[0]);

  it.each([
    'Which are the 10 largest debits in Bank Statement?',
    'What period does this statement cover?',
    'What was the lowest balance and on which date?',
    'How much did I spend at FUEL STATION?',
  ])('routes "%s" to the data lane with no model call', async (message) => {
    const decision = await route(message, 'action');
    expect(decision.route).toBe('data');
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('still treats a request for a change as a write', async () => {
    const decision = await route('Highlight every row where Debit is more than 10000', 'action');
    expect(decision.route).toBe('write');
  });

  it('leaves the same question on the ask route in Ask mode, where it is computed too', async () => {
    const decision = await route('Which are the 10 largest debits in Bank Statement?', 'ask');
    expect(decision.route).toBe('ask');
    expect(openRouter.complete).not.toHaveBeenCalled();
  });
});
