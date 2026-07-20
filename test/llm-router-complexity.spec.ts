import { LlmRouterService } from '../src/excel-ai/services/llm-router.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { RouterInput } from '../src/excel-ai/types/router.types';

describe('LlmRouterService complexity integration', () => {
  const baseInput: RouterInput = {
    message: '',
    mode: 'action',
    sheetHeaders: ['Amount', 'Status'],
    activeSheet: 'Sheet1',
  };

  let openRouter: jest.Mocked<Pick<OpenRouterService, 'complete'>>;
  let config: jest.Mocked<Pick<AppConfigService, never>>;
  let service: LlmRouterService;

  beforeEach(() => {
    openRouter = {
      complete: jest.fn(),
    };
    config = {} as jest.Mocked<Pick<AppConfigService, never>>;
    service = new LlmRouterService(openRouter as unknown as OpenRouterService, config as AppConfigService);
  });

  it('returns write route with regex-matched complexity', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'bold cells A1 to C1',
    });

    expect(decision).toEqual({
      route: 'write',
      complexity: 0,
      actionHint: 'CELL_FORMAT',
      matchedBy: 'regex',
      confidence: 1.0,
      reasoning: 'Complexity regex: tier=0 hint=CELL_FORMAT',
    });
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('returns compound write as tier 3 via regex', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'sort by column B and then create a chart',
    });

    expect(decision.route).toBe('write');
    expect(decision.complexity).toBe(3);
    expect(decision.matchedBy).toBe('regex');
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('does not add complexity for shortcut route', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'freeze top row',
    });

    expect(decision.route).toBe('shortcut');
    expect(decision.complexity).toBeUndefined();
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('does not add complexity for data route', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'sum column B',
    });

    expect(decision.route).toBe('data');
    expect(decision.complexity).toBeUndefined();
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('write-intent-guard overrides ask route in non-action mode when message is a mutation', async () => {
    const decision = await service.route({
      ...baseInput,
      mode: 'ask',
      message: 'bold cells A1 to C1',
    });

    expect(decision.route).toBe('write');
    expect(decision.overridden).toBe(true);
    expect(decision.complexity).toBe(0);
    expect(decision.actionHint).toBe('CELL_FORMAT');
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('keeps ask route in non-action mode for genuine questions', async () => {
    const decision = await service.route({
      ...baseInput,
      mode: 'ask',
      message: 'what does this workbook contain',
    });

    expect(decision.route).toBe('ask');
    expect(decision.complexity).toBeUndefined();
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('defaults write complexity to tier 3 when LLM omits it', async () => {
    openRouter.complete.mockResolvedValue(
      JSON.stringify({
        route: 'write',
        confidence: 0.82,
        reasoning: 'User wants to rename a sheet',
      }),
    );

    const decision = await service.route({
      ...baseInput,
      message: 'rename this sheet to Q1 Summary',
    });

    expect(decision.route).toBe('write');
    expect(decision.complexity).toBe(3);
    expect(decision.matchedBy).toBe('llm-fallback');
    expect(openRouter.complete).toHaveBeenCalledTimes(1);
  });

  it('uses LLM-provided complexity for write routes', async () => {
    openRouter.complete.mockResolvedValue(
      JSON.stringify({
        route: 'write',
        complexity: 2,
        confidence: 0.9,
        reasoning: 'Single formula write',
        actionHint: 'FORMULA_GEN',
      }),
    );

    const decision = await service.route({
      ...baseInput,
      message: 'add margin column based on revenue and cost',
    });

    expect(decision.route).toBe('write');
    expect(decision.complexity).toBe(2);
    expect(decision.matchedBy).toBe('llm-fallback');
  });

  it('does not add complexity for non-write LLM routes', async () => {
    openRouter.complete.mockResolvedValue(
      JSON.stringify({
        route: 'export',
        confidence: 0.88,
        reasoning: 'Find and copy matching rows',
      }),
    );

    const decision = await service.route({
      ...baseInput,
      message: 'copy all paid invoices to a new sheet',
    });

    expect(decision.route).toBe('export');
    expect(decision.complexity).toBeUndefined();
  });

  it('adds complexity 3 on fallback write decisions', async () => {
    openRouter.complete.mockRejectedValue(new Error('router down'));

    const decision = await service.route({
      ...baseInput,
      message: 'format column C as currency',
    });

    expect(decision.route).toBe('write');
    expect(decision.complexity).toBe(3);
    expect(decision.matchedBy).toBe('llm-fallback');
  });
});
