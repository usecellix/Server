import { LlmRouterService } from '../src/excel-ai/services/llm-router.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { RouterInput } from '../src/excel-ai/types/router.types';

describe('LlmRouterService write-intent guard', () => {
  const baseInput: RouterInput = {
    message: '',
    mode: 'action',
    sheetHeaders: ['Invoice', 'Total Amount', 'Status'],
    activeSheet: 'Sheet1',
  };

  let openRouter: jest.Mocked<Pick<OpenRouterService, 'complete'>>;
  let service: LlmRouterService;

  beforeEach(() => {
    openRouter = {
      complete: jest.fn(),
    };
    service = new LlmRouterService(
      openRouter as unknown as OpenRouterService,
      {} as AppConfigService,
    );
  });

  const repro = 'sort the sheet based on Total Amount descending';

  it('routes sort+Total Amount to write (never data census)', async () => {
    const decision = await service.route({ ...baseInput, message: repro });

    expect(decision.route).toBe('write');
    expect(decision.route).not.toBe('data');
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('routes the repro to write on every of 10 consecutive calls (no flapping)', async () => {
    for (let i = 0; i < 10; i += 1) {
      const decision = await service.route({ ...baseInput, message: repro });
      expect(decision.route).toBe('write');
    }
  });

  it('routes sort by total + highlight to write', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'sort by total and then highlight the top row',
    });

    expect(decision.route).toBe('write');
  });

  it('does not override pure data queries', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'sum column Total Amount',
    });

    expect(decision.route).toBe('data');
    expect(decision.overridden).toBeUndefined();
  });

  it('routes soft multi-month + dashboard scaffold to write (not data)', async () => {
    const message =
      'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of the remaining sheets, in the main sheet i need to have dashboard also, my need to record payments and related things ,which all month sheets include Unit No, Guest, Guest name, check in, check out, Rate per night, total amount, source, payment status, bank account';
    const decision = await service.route({ ...baseInput, message });

    expect(decision.route).toBe('write');
    expect(decision.complexity).toBe(3);
    expect(decision.route).not.toBe('data');
    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('does not override instant shortcuts', async () => {
    const decision = await service.route({
      ...baseInput,
      message: 'freeze top row',
    });

    expect(decision.route).toBe('shortcut');
    expect(decision.overridden).toBeUndefined();
  });

  it('overrides ask-mode short-circuit when message has write intent', async () => {
    const decision = await service.route({
      ...baseInput,
      mode: 'ask',
      message: 'bold cells A1 to C1',
    });

    expect(decision.route).toBe('write');
    expect(decision.overridden).toBe(true);
    expect(decision.complexity).toBe(0);
    expect(decision.actionHint).toBe('CELL_FORMAT');
  });

  it('does not override ask-mode for genuine questions', async () => {
    const decision = await service.route({
      ...baseInput,
      mode: 'ask',
      message: 'what is in this workbook',
    });

    expect(decision.route).toBe('ask');
    expect(decision.overridden).toBeUndefined();
  });
});
