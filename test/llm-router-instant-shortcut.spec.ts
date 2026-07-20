import { LlmRouterService } from '../src/excel-ai/services/llm-router.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';

describe('LlmRouterService.peekInstantShortcut (Spec 09 item 3)', () => {
  let router: LlmRouterService;

  beforeEach(() => {
    router = new LlmRouterService(
      {} as OpenRouterService,
      {} as AppConfigService,
    );
  });

  it('detects freeze/zoom/protect without calling the LLM', () => {
    expect(router.peekInstantShortcut('freeze top row')?.action).toBe('FREEZE_PANES');
    expect(router.peekInstantShortcut('zoom to 150%')?.action).toBe('SET_ZOOM');
    expect(router.peekInstantShortcut('protect this sheet')?.action).toBe('PROTECT_SHEET');
  });

  it('returns null for non-shortcut messages', () => {
    expect(router.peekInstantShortcut('lookup the value techpath786@gmail.com')).toBeNull();
    expect(router.peekInstantShortcut('what is the total CGST')).toBeNull();
  });
});

describe('LlmRouterService.find+export routing', () => {
  let router: LlmRouterService;

  beforeEach(() => {
    router = new LlmRouterService(
      {} as OpenRouterService,
      {} as AppConfigService,
    );
  });

  it('routes find + copy to new sheet as export, not read-only data', async () => {
    const decision = await router.route({
      message: 'find Applied and copy those rows to a new sheet',
      mode: 'action',
      activeSheet: 'DemoApps',
      sheetHeaders: ['Job Title', 'Company', 'Student Name', 'Student Email', 'Status'],
    });
    expect(decision.route).toBe('export');
  });

  it('still routes pure find as data', async () => {
    const decision = await router.route({
      message: 'find Applied',
      mode: 'action',
      activeSheet: 'DemoApps',
      sheetHeaders: ['Status'],
    });
    expect(decision.route).toBe('data');
  });
});
