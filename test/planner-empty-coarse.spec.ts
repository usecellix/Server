import { PlannerAgent } from '../src/agents/planner.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #335 — live: the coarse pass answered `{"phases":[]}` in 940ms. It
 * was accepted as "Planned 0 steps", every check passed an empty build, and
 * the user saw "Something went wrong applying this change — try rephrasing".
 */
const PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of the ' +
  'remaining sheets, in the main sheet i need to have dashboard also, my need to record payments and related ' +
  'things ,which all month sheets include Unit No, Guest, Guest name, check in, check out, Rate per night, ' +
  'total amount, source, payment status, bank account';
const context: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [{
    name: 'Sheet1', usedRange: 'A1', rowCount: 0, columnCount: 0, values: [], formulas: [],
    numberFormats: [], structure: 'unknown', headerRowIndex: 0,
  }],
  namedRanges: [],
  tables: [],
};
const EMPTY = JSON.stringify({ phases: [], clarificationsNeeded: [], confidence: 'medium', reasoning: '' });
const COARSE = JSON.stringify({
  phases: [{ id: 'p1', kind: 'Create the Lists sheet', targetSheet: 'Lists', dependsOn: [] }],
  clarificationsNeeded: [], confidence: 'high', reasoning: 'ok',
});
const PHASE = JSON.stringify({
  subtasks: [{ id: 's1', description: "Create sheet 'Lists' with Source, Payment Status", targetSheet: 'Lists', dependsOn: [], estimatedActions: 3 }],
  clarificationsNeeded: [], confidence: 'high', reasoning: 'ok',
});

function agent(responses: string[]) {
  const complete = jest.fn();
  for (const r of responses) complete.mockResolvedValueOnce(r);
  complete.mockResolvedValue(PHASE);
  const llm = { complete } as unknown as OpenRouterService;
  const config = { openRouterModelHigh: 'm', openRouterModelPlanner: 'm' } as unknown as AppConfigService;
  return { planner: new PlannerAgent(llm, config), complete };
}

describe('PlannerAgent — an empty coarse plan is a failed answer (TASKS.md #335)', () => {
  it('retries an empty phase list instead of planning zero steps', async () => {
    const { planner, complete } = agent([EMPTY, COARSE, PHASE]);
    const plan = await planner.plan(PROMPT, context, [], undefined, 'c', undefined, 3);
    expect(complete.mock.calls[1][0].userMessage).toBeDefined(); // the retry happened
    expect(plan.subtasks.length).toBeGreaterThan(0);
  });

  it('falls back to one all-covering phase when both attempts are empty — never zero steps', async () => {
    const { planner } = agent([EMPTY, EMPTY, PHASE]);
    const plan = await planner.plan(PROMPT, context, [], undefined, 'c', undefined, 3);
    expect(plan.subtasks.length).toBeGreaterThan(0);
  });

  it('still returns a blocking clarification with no phases as-is', async () => {
    const ask = JSON.stringify({ phases: [], clarificationsNeeded: ['Which year?'], confidence: 'low', reasoning: '' });
    const { planner, complete } = agent([ask]);
    const plan = await planner.plan(PROMPT, context, [], undefined, 'c', undefined, 3);
    expect(plan.clarificationsNeeded).toEqual(['Which year?']);
    expect(complete).toHaveBeenCalledTimes(1);
  });
});
