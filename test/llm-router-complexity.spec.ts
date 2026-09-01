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

  it('escalates an unclassified write to tier 3 without consulting the LLM', async () => {
    // "rename" is a WRITE_INTENT_VERB that classifyComplexity does not recognize,
    // so the pre-LLM write-intent guard decides the route by regex alone.
    const decision = await service.route({
      ...baseInput,
      message: 'rename this sheet to Q1 Summary',
    });

    expect(decision.route).toBe('write');
    expect(decision.complexity).toBe(3);
    expect(decision.matchedBy).toBe('regex');
    expect(decision.overridden).toBe(true);
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('escalates an unrecognized formula-style write to tier 3 by regex', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'add margin column based on revenue and cost',
    });

    expect(decision.route).toBe('write');
    expect(decision.complexity).toBe(3);
    expect(decision.matchedBy).toBe('regex');
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('routes a copy-to-new-sheet request to write rather than a read-only lane', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'copy all paid invoices to a new sheet',
    });

    expect(decision.route).toBe('write');
    expect(decision.overridden).toBe(true);
  });

  it('Spec 20 repro — create sheet + copy paid rows routes to write', async () => {
    const decision = await service.route({
      ...baseInput,
      message:
        'create a new sheet named Paid payments and copy the paid data from purchase register to that new sheet only paid',
    });

    expect(decision.route).toBe('write');
    expect(decision.overridden).toBe(true);
  });

  // In action mode the pre-LLM write-intent guard intercepts every verb that
  // fallbackDecision would call a write, so the LLM-failure path can no longer
  // produce one. The write-intent override in non-action modes is what still can.
  it('adds complexity 3 when a non-action mode request is overridden to write', async () => {
    const decision = await service.route({
      ...baseInput,
      mode: 'ask',
      message: 'format column C as currency',
    });

    expect(decision.route).toBe('write');
    expect(decision.complexity).toBe(3);
    expect(decision.matchedBy).toBe('llm-fallback');
    expect(decision.overridden).toBe(true);
  });
});
