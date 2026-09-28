import { PlannerAgent } from '../src/agents/planner.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Live incident (TASKS.md #282): a turn resuming after a clarifying question
 * ("dd-mm-yyyy") produced a plan covering only Main + Lists + ONE month,
 * instead of the full 12-month build the ORIGINAL prompt (now living in
 * `history`, per TASKS.md #259) asked for. Traced to `needsTwoPassPlanning`
 * and `resolvePlannerMaxTokens` both scoring the 10-character reply alone —
 * zero object keywords, under the length threshold — so a compound build
 * that would normally get two-pass planning and a large token budget got
 * neither, on exactly the turn continuing that same compound build.
 */

const LONG_ORIGINAL_PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of the ' +
  'remaining sheets, in the main sheet i need to have dashboard also, my need to record payments and related things ' +
  ',which all month sheets include Unit No, Guest, Guest name, check in, check out, Rate per night, total amount, ' +
  'source, payment status, bank account';

const context: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [{
    name: 'Sheet1', usedRange: 'A1:A1', rowCount: 1, columnCount: 1,
    values: [['']], formulas: [['']], numberFormats: [['General']],
    structure: 'data_table', headerRowIndex: 0,
  }],
  namedRanges: [],
  tables: [],
};

const simplePlanJson = JSON.stringify({
  subtasks: [{ id: 's1', description: 'x', targetSheet: 'Sheet1', dependsOn: [], estimatedActions: 1 }],
  clarificationsNeeded: [], confidence: 'high', reasoning: '',
});

function buildAgent(completeImpl: jest.Mock): PlannerAgent {
  const llm = { complete: completeImpl } as unknown as OpenRouterService;
  const config = {
    openRouterModelHigh: 'openai/gpt-5',
    openRouterModelPlanner: 'openai/gpt-5',
  } as unknown as AppConfigService;
  return new PlannerAgent(llm, config);
}

describe('PlannerAgent — sizing sees the resumed conversation, not just the reply (TASKS.md #282)', () => {
  it('a short reply ALONE (no history) stays single-pass — no behavior change for the common case', async () => {
    const complete = jest.fn().mockResolvedValue(simplePlanJson);
    const agent = buildAgent(complete);

    await agent.plan('dd-mm-yyyy', context, []);

    expect(complete).toHaveBeenCalledTimes(1); // single-pass: exactly one call
  });

  it('the SAME short reply, with the original long request in history, routes to two-pass planning', async () => {
    // Coarse pass, then one phase-expansion call — planTwoPass's own shape.
    const complete = jest
      .fn()
      .mockResolvedValueOnce(
        JSON.stringify({
          phases: [{ id: 'p1', kind: 'Build Main', targetSheet: 'Main', dependsOn: [] }],
          clarificationsNeeded: [], confidence: 'high', reasoning: '',
        }),
      )
      .mockResolvedValue(
        JSON.stringify({
          subtasks: [{ id: 's1', description: 'Create Main', targetSheet: 'Main', dependsOn: [], estimatedActions: 3 }],
          clarificationsNeeded: [], confidence: 'high', reasoning: '',
        }),
      );
    const agent = buildAgent(complete);

    await agent.plan('dd-mm-yyyy', context, [
      { role: 'user', content: LONG_ORIGINAL_PROMPT },
      { role: 'assistant', content: 'What date format do you want?' },
    ]);

    // planTwoPass makes at least 2 calls (coarse + >=1 phase expansion);
    // single-pass planning makes exactly 1. This alone proves two-pass fired.
    expect(complete.mock.calls.length).toBeGreaterThanOrEqual(2);
  });

  it('an assistant turn is never folded into the sizing signal, only prior USER turns', async () => {
    const complete = jest.fn().mockResolvedValue(simplePlanJson);
    const agent = buildAgent(complete);

    // The long text sits on the ASSISTANT side only — must not trigger two-pass.
    await agent.plan('dd-mm-yyyy', context, [
      { role: 'assistant', content: LONG_ORIGINAL_PROMPT },
    ]);

    expect(complete).toHaveBeenCalledTimes(1); // single-pass: exactly one call
  });
});
