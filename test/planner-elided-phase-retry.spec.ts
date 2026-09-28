import { PlannerAgent } from '../src/agents/planner.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { WorkbookContext } from '../src/agents/types/agent.types';
import { dropElidedRepeatEntries, isElidedBatch } from '../src/agents/utils/plan-coverage.util';

/**
 * TASKS.md #325 — live: the month phase of a two-pass plan came back as
 * stubs, and the plan-wide #271 check threw the WHOLE plan away: "Planner
 * returned 12 elided subtask descriptions out of 20", nothing built, 14s in.
 * One phase eliding must cost a retry of that phase, not the build.
 */

const PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of the ' +
  'remaining sheets, in the main sheet i need to have dashboard also, my need to record payments and related ' +
  'things ,which all month sheets include Unit No, Guest, Guest name, check in, check out, Rate per night, ' +
  'total amount, source, payment status, bank account';

const MONTHS = ['January', 'February', 'March'];

const context: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1',
      rowCount: 0,
      columnCount: 0,
      values: [],
      formulas: [],
      numberFormats: [],
      structure: 'unknown',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

const coarse = JSON.stringify({
  phases: [
    { id: 'p1', kind: 'Create one sheet per month', targetSheet: 'January', dependsOn: [], repeatFor: MONTHS },
    { id: 'p2', kind: 'Build the Main dashboard', targetSheet: 'Main', dependsOn: ['p1'] },
  ],
  clarificationsNeeded: [],
  confidence: 'high',
  reasoning: 'months then dashboard',
});

const whole = (month: string) =>
  `Create sheet '${month}' with headers Unit No | Guest | Guest Name | Check In | Check Out | Rate Per Night | ` +
  `Total Amount | Source | Payment Status | Bank Account and table tbl${month} over A1:J2.`;

function phase(subtasks: Array<{ id: string; description: string; targetSheet: string }>) {
  return JSON.stringify({
    subtasks: subtasks.map((s) => ({ ...s, dependsOn: [], estimatedActions: 4 })),
    clarificationsNeeded: [],
    confidence: 'high',
    reasoning: 'ok',
  });
}

const stubMonths = phase(MONTHS.map((m, i) => ({ id: `s${i + 1}`, description: 'Create sheet...', targetSheet: m })));
const wholeMonths = phase(MONTHS.map((m, i) => ({ id: `s${i + 1}`, description: whole(m), targetSheet: m })));
const main = phase([{ id: 's1', description: 'Write the Main title and KPI row A1:F2 summing the month sheets.', targetSheet: 'Main' }]);

function agentWith(responses: string[]): { agent: PlannerAgent; complete: jest.Mock } {
  const complete = jest.fn();
  for (const r of responses) complete.mockResolvedValueOnce(r);
  const llm = { complete } as unknown as OpenRouterService;
  const config = { openRouterModelHigh: 'm', openRouterModelPlanner: 'm' } as unknown as AppConfigService;
  return { agent: new PlannerAgent(llm, config), complete };
}

const plan = (agent: PlannerAgent) => agent.plan(PROMPT, context, [], undefined, 'corr', undefined, 3);

describe('PlannerAgent — an elided phase is re-asked, not fatal (TASKS.md #325)', () => {
  it('re-expands a stubbed phase once and uses the whole answer', async () => {
    const { agent, complete } = agentWith([coarse, stubMonths, wholeMonths, main]);

    const result = await plan(agent);

    expect(complete).toHaveBeenCalledTimes(4);
    // The re-ask says what went wrong.
    expect(complete.mock.calls[2][0].userMessage).toMatch(/shortened subtask descriptions into stubs/);
    const months = result.subtasks.filter((s) => MONTHS.includes(s.targetSheet));
    expect(months).toHaveLength(3);
    expect(months.every((s) => s.description.includes('Unit No | Guest'))).toBe(true);
  });

  it('rebuilds entries still elided after the retry from a sibling that came back whole', async () => {
    const partlyWhole = phase([
      { id: 's1', description: whole('January'), targetSheet: 'January' },
      { id: 's2', description: 'Create sheet...', targetSheet: 'February' },
      { id: 's3', description: 'Create sheet...', targetSheet: 'March' },
    ]);
    const { agent } = agentWith([coarse, stubMonths, partlyWhole, main]);

    const result = await plan(agent);

    const march = result.subtasks.find((s) => s.targetSheet === 'March');
    expect(march?.description).toBe(whole('March'));
    expect(result.subtasks.filter((s) => MONTHS.includes(s.targetSheet))).toHaveLength(3);
  });

  it('still refuses the plan when every attempt is stubs — nothing to build a faithful sheet from', async () => {
    const { agent } = agentWith([coarse, stubMonths, stubMonths, main]);
    await expect(plan(agent)).rejects.toThrow(/elided subtask descriptions/);
  });
});

describe('elision helpers', () => {
  it('treats a batch as elided only when stubs are a real share of it', () => {
    const s = (description: string) => ({ id: 'x', description, targetSheet: 'A', dependsOn: [], estimatedActions: 1 });
    expect(isElidedBatch([s('Create sheet...'), s('Create sheet...')])).toBe(true);
    expect(isElidedBatch([s('Hide the Lists sheet'), s(whole('January'))])).toBe(false);
    expect(isElidedBatch([s('Set formulas...'), s(whole('A')), s(whole('B')), s(whole('C'))])).toBe(false);
  });

  it('never drops entries when no sibling came back whole', () => {
    const p = { id: 'p1', kind: 'k', targetSheet: 'January', dependsOn: [], repeatFor: MONTHS };
    const stubs = MONTHS.map((m) => ({ id: m, description: 'Create sheet...', targetSheet: m, dependsOn: [], estimatedActions: 1 }));
    expect(dropElidedRepeatEntries(p, stubs)).toEqual({ subtasks: stubs, dropped: [] });
  });
});
