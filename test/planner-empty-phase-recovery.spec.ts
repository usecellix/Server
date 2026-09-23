import { PlannerAgent } from '../src/agents/planner.agent';
import { synthesizeSubtasksForEmptyPhase } from '../src/agents/utils/plan-coverage.util';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { PlanPhase, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Live incident (TASKS.md #285): the coarse pass correctly identified three
 * phases — Lists, the twelve month sheets (`repeatFor` Jan–Dec), and Main —
 * but phase p2's expansion came back empty. `expandPhase` logged it, returned
 * an empty plan, and the build carried on: 6 subtasks instead of 18, "Step 1
 * of 3 … ✓ Applied", and a Main dashboard summing twelve sheets that were
 * never created. `ensureRepeatForCoverage` could not save it — it clones a
 * SIBLING subtask, and an empty expansion has no sibling to clone.
 *
 * Confirmed straight from planner.log for that run:
 *   parsed subtasks: 6 -> p1_s1,p3_s1,p3_s2,p3_s3,p3_s4,p3_s5   (no p2_* at all)
 */

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

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

/** The real p2 phase from the live run. */
const MONTHS_PHASE: PlanPhase = {
  id: 'p2',
  kind:
    'Create one sheet per month with the standard booking/payment columns (Unit No, Guest, Guest name, ' +
    'Check in, Check out, Rate per night, Total amount, Source, Payment Status, Bank Account) with ' +
    'dropdowns referencing the Lists sheet',
  targetSheet: 'January',
  dependsOn: ['p1'],
  repeatFor: MONTHS,
};

describe('synthesizeSubtasksForEmptyPhase (TASKS.md #285)', () => {
  it('covers every repeatFor entry — all twelve months, one subtask each', () => {
    const subtasks = synthesizeSubtasksForEmptyPhase(MONTHS_PHASE);
    expect(subtasks).toHaveLength(12);
    expect(subtasks.map((s) => s.targetSheet)).toEqual(MONTHS);
    expect(subtasks.map((s) => s.id)).toEqual(MONTHS.map((_, i) => `s${i + 1}`));
  });

  it("carries the phase's own description into each subtask, so the build knows what to make", () => {
    const [january] = synthesizeSubtasksForEmptyPhase(MONTHS_PHASE);
    expect(january.description).toContain("Create sheet 'January'");
    expect(january.description).toContain('booking/payment columns');
  });

  it('falls back to the single targetSheet for a phase with no repeatFor', () => {
    const subtasks = synthesizeSubtasksForEmptyPhase({
      id: 'p3', kind: 'Build the Main dashboard', targetSheet: 'Main', dependsOn: ['p1', 'p2'],
    });
    expect(subtasks).toHaveLength(1);
    expect(subtasks[0].targetSheet).toBe('Main');
  });

  it('yields nothing when the phase carries neither repeatFor nor a target sheet', () => {
    expect(
      synthesizeSubtasksForEmptyPhase({ id: 'p9', kind: 'something', targetSheet: '', dependsOn: [] }),
    ).toEqual([]);
  });
});

describe('PlannerAgent two-pass — an empty phase never silently vanishes (TASKS.md #285)', () => {
  function buildAgent(completeImpl: jest.Mock): PlannerAgent {
    const llm = { complete: completeImpl } as unknown as OpenRouterService;
    const config = {
      openRouterModelHigh: 'openai/gpt-5',
      openRouterModelPlanner: 'openai/gpt-5',
    } as unknown as AppConfigService;
    return new PlannerAgent(llm, config);
  }

  const COARSE = JSON.stringify({
    phases: [
      { id: 'p1', kind: 'Create a Lists sheet', targetSheet: 'Lists', dependsOn: [] },
      { ...MONTHS_PHASE },
      { id: 'p3', kind: 'Build the Main sheet', targetSheet: 'Main', dependsOn: ['p1', 'p2'] },
    ],
    clarificationsNeeded: [],
    confidence: 'high',
    reasoning: '',
  });

  const phasePlan = (id: string, sheet: string) =>
    JSON.stringify({
      subtasks: [{ id: 's1', description: `Build ${sheet}`, targetSheet: sheet, dependsOn: [], estimatedActions: 3 }],
      clarificationsNeeded: [], confidence: 'high', reasoning: '',
    });

  const LONG_PROMPT =
    'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details ' +
    'of the remaining sheets, in the main sheet i need to have dashboard also, my need to record payments ' +
    'and related things, which all month sheets include Unit No, Guest, Guest name, check in, check out, ' +
    'Rate per night, total amount, source, payment status, bank account';

  it('recovers all twelve months when the months phase expands to nothing (the live shape)', async () => {
    // Coarse pass, then: p1 fine, p2 UNPARSEABLE both times (the live failure),
    // p3 fine. `expandPhase` retries once before giving up, hence two duds.
    const complete = jest
      .fn()
      .mockResolvedValueOnce(COARSE)
      .mockResolvedValueOnce(phasePlan('p1', 'Lists'))
      .mockResolvedValueOnce('not json at all')
      .mockResolvedValueOnce('still not json')
      .mockResolvedValue(phasePlan('p3', 'Main'));

    const plan = await buildAgent(complete).plan(LONG_PROMPT, context, []);

    const monthSubtasks = plan.subtasks.filter((s) => MONTHS.includes(s.targetSheet));
    expect(monthSubtasks).toHaveLength(12);
    expect(new Set(monthSubtasks.map((s) => s.targetSheet))).toEqual(new Set(MONTHS));
    // Recovered, so the plan must NOT claim it couldn't be planned...
    expect(plan.clarificationsNeeded.join(' ')).not.toContain('could not plan');
    // ...and must not block the build with a low-confidence gate either.
    expect(plan.confidence).not.toBe('low');
  });

  it('every recovered month subtask is wired to the phase it depends on', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(COARSE)
      .mockResolvedValueOnce(phasePlan('p1', 'Lists'))
      .mockResolvedValueOnce('not json at all')
      .mockResolvedValueOnce('still not json')
      .mockResolvedValue(phasePlan('p3', 'Main'));

    const plan = await buildAgent(complete).plan(LONG_PROMPT, context, []);

    const listsId = plan.subtasks.find((s) => s.targetSheet === 'Lists')?.id;
    const february = plan.subtasks.find((s) => s.targetSheet === 'February');
    expect(listsId).toBeDefined();
    expect(february?.dependsOn).toContain(listsId);
  });

  it('a fully successful two-pass plan is completely unaffected', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(COARSE)
      .mockResolvedValueOnce(phasePlan('p1', 'Lists'))
      .mockResolvedValueOnce(
        JSON.stringify({
          subtasks: MONTHS.map((m, i) => ({
            id: `s${i + 1}`, description: `Build ${m}`, targetSheet: m, dependsOn: [], estimatedActions: 3,
          })),
          clarificationsNeeded: [], confidence: 'high', reasoning: '',
        }),
      )
      .mockResolvedValue(phasePlan('p3', 'Main'));

    const plan = await buildAgent(complete).plan(LONG_PROMPT, context, []);

    expect(plan.confidence).toBe('high');
    expect(plan.subtasks.filter((s) => MONTHS.includes(s.targetSheet))).toHaveLength(12);
  });
});
