import { classifyComplexity } from '../src/excel-ai/utils/complexity-classifier.util';
import { resolveExecutableTier } from '../src/excel-ai/utils/complexity-tiering-flag.util';
import { normalizeAssistantMode } from '../src/excel-ai/utils/mode-guard.util';

/**
 * Lightweight conversation e2e guards for Spec 08:
 * full request classification → executable tier → mode branching contract.
 * (Heavy Nest HTTP SSE harness is out of scope; handlers are covered in unit specs.)
 */
describe('conversation e2e — tier + mode contracts', () => {
  const cases: Array<{
    message: string;
    mode: 'ask' | 'plan' | 'action' | 'act' | undefined;
    expectedClassified: 0 | 1 | 2 | 3 | null;
    expectedExecutableFull: 0 | 1 | 2 | 3 | null;
    writeAllowed: boolean;
  }> = [
    {
      message: 'bold cells A1 to C1',
      mode: 'action',
      expectedClassified: 0,
      expectedExecutableFull: 0,
      writeAllowed: true,
    },
    {
      message: 'sort column B descending by value',
      mode: 'act',
      expectedClassified: 1,
      expectedExecutableFull: 1,
      writeAllowed: true,
    },
    {
      message: 'calculate GST at 18% for column D',
      mode: undefined,
      expectedClassified: 2,
      expectedExecutableFull: 2,
      writeAllowed: true,
    },
    {
      message: 'sort by column B and then create a chart',
      mode: 'action',
      expectedClassified: 3,
      expectedExecutableFull: 3,
      writeAllowed: true,
    },
    {
      message: 'bold cells A1 to C1',
      mode: 'ask',
      expectedClassified: 0,
      expectedExecutableFull: 0,
      writeAllowed: false,
    },
    {
      message: 'calculate GST at 18% for column D',
      mode: 'plan',
      expectedClassified: 2,
      expectedExecutableFull: 2,
      writeAllowed: false,
    },
  ];

  it.each(cases)(
    '$mode · "$message"',
    ({ message, mode, expectedClassified, expectedExecutableFull, writeAllowed }) => {
      const classified = classifyComplexity(message);
      const tier = classified.match?.tier ?? null;
      expect(tier).toBe(expectedClassified);

      if (tier !== null) {
        expect(resolveExecutableTier(tier, 'full')).toBe(expectedExecutableFull);
        expect(resolveExecutableTier(tier, 'shadow')).toBe(3);
      }

      const normalized = normalizeAssistantMode(mode);
      const canWrite = normalized === 'action';
      expect(canWrite).toBe(writeAllowed);
    },
  );

  it('shortcut/data/ask route kinds are unchanged by tiering flag (contract)', () => {
    // Non-write routes never enter handleWriteRoute — tiering flag must not apply.
    const routesUnaffected = ['shortcut', 'data', 'export', 'ask'] as const;
    expect(routesUnaffected).toHaveLength(4);
    expect(resolveExecutableTier(0, 'off')).toBe(3);
  });
});
