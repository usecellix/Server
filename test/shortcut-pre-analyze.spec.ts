import { routeShortcutAction } from '../src/excel-ai/utils/shortcut-router.util';
import { LlmRouterService } from '../src/excel-ai/services/llm-router.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';

/**
 * Spec 09 item 3: instant shortcuts are fully resolvable without SheetAnalyzer.
 * ConversationService runs peekInstantShortcut + routeShortcutAction before analyze().
 */
describe('instant shortcut pre-analyze contract (Spec 09 item 3)', () => {
  const router = new LlmRouterService({} as OpenRouterService, {} as AppConfigService);
  const analyze = jest.fn();

  beforeEach(() => {
    analyze.mockClear();
  });

  it('resolves freeze top row without invoking sheet analysis', () => {
    const message = 'freeze top row';
    const peek = router.peekInstantShortcut(message);
    expect(peek?.action).toBe('FREEZE_PANES');

    const actions = routeShortcutAction(message, 'Sheet1');
    expect(actions?.some((a) => a.type === 'FREEZE_PANES')).toBe(true);
    expect(analyze).not.toHaveBeenCalled();
  });
});
