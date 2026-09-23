import { checkPlanIntegrity } from '../src/agents/utils/plan-integrity.util';
import { PlannerOutput, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 6 acceptance:
 *  - "The #285 run shape (p2 empty) is caught at plan time, not from the workbook."
 *  - "A plan referencing an unplanned sheet fails the gate."
 *  - "A correct 12-month plan passes with zero added latency and no model call."
 */

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details ' +
  'of the remaining sheets, in the main sheet i need to have dashboard also, my need to record payments';

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

const monthSubtask = (month: string, i: number): SubTask => ({
  id: `p2_s${i + 1}`,
  description: `Create sheet '${month}' and build the booking log with headers and formulas.`,
  targetSheet: month,
  dependsOn: ['p1_s1'],
  estimatedActions: 12,
});

const planOf = (subtasks: SubTask[]): PlannerOutput => ({
  subtasks,
  clarificationsNeeded: [],
  confidence: 'high',
  reasoning: '',
});

const lists: SubTask = {
  id: 'p1_s1', description: "Create sheet 'Lists' with lookup values", targetSheet: 'Lists',
  dependsOn: [], estimatedActions: 4,
};

describe('checkPlanIntegrity (Phase 6, TASKS.md #291)', () => {
  it('passes a correct 12-month plan untouched', () => {
    const plan = planOf([lists, ...MONTHS.map(monthSubtask)]);
    const result = checkPlanIntegrity({ prompt: PROMPT, plan, context });

    expect(result.violations).toEqual([]);
    expect(result.repaired).toEqual([]);
    expect(result.plan.subtasks).toHaveLength(13);
  });

  it('catches the #285 shape — the request asks for every month, the plan covers three', () => {
    const plan = planOf([lists, ...MONTHS.slice(0, 3).map(monthSubtask)]);
    const result = checkPlanIntegrity({ prompt: PROMPT, plan, context });

    const violation = result.violations.find((v) => v.kind === 'missing-repeated-entity');
    expect(violation).toBeDefined();
    expect(violation!.detail).toContain('3 of 12');

    // ...and repairs it from the months that WERE planned, so the build still runs.
    expect(result.repaired).toHaveLength(9);
    expect(new Set(result.plan.subtasks.map((s) => s.targetSheet))).toEqual(
      new Set(['Lists', ...MONTHS]),
    );
  });

  it('re-targets a cloned month properly rather than leaving the template’s name in it', () => {
    const plan = planOf([lists, monthSubtask('January', 0)]);
    const { plan: repaired } = checkPlanIntegrity({ prompt: PROMPT, plan, context });

    const july = repaired.subtasks.find((s) => s.targetSheet === 'July')!;
    expect(july.description).toContain("'July'");
    expect(july.description).not.toContain('January');
  });

  it('is FATAL when the request wants every month and the plan has none to model them on', () => {
    // Exactly the live #285 outcome: Lists + Main only, no month anywhere.
    const main: SubTask = {
      id: 'p3_s1', description: 'Create the Main dashboard', targetSheet: 'Main',
      dependsOn: [], estimatedActions: 8,
    };
    const result = checkPlanIntegrity({ prompt: PROMPT, plan: planOf([lists, main]), context });

    const violation = result.violations.find((v) => v.kind === 'missing-repeated-entity');
    expect(violation?.fatal).toBe(true);
    expect(result.repaired).toEqual([]);
  });

  it('does NOT invent months for a request that only asks about one', () => {
    const plan = planOf([monthSubtask('January', 0)]);
    const result = checkPlanIntegrity({
      prompt: 'Add a January sheet for my bookings',
      plan,
      context,
    });
    expect(result.violations.filter((v) => v.kind === 'missing-repeated-entity')).toEqual([]);
    expect(result.repaired).toEqual([]);
  });

  it('reports a dependency that names no step in the plan', () => {
    const orphan: SubTask = {
      id: 's9', description: 'Write totals', targetSheet: 'Lists',
      dependsOn: ['does_not_exist'], estimatedActions: 2,
    };
    const result = checkPlanIntegrity({
      prompt: 'Write totals on Lists',
      plan: planOf([lists, orphan]),
      context,
    });
    expect(result.violations.some((v) => v.kind === 'dangling-dependency')).toBe(true);
  });

  it('reports a step targeting a sheet nothing creates and that does not already exist', () => {
    const writer: SubTask = {
      id: 's5', description: 'Write the summary totals', targetSheet: 'Summary',
      dependsOn: [], estimatedActions: 3,
    };
    const result = checkPlanIntegrity({
      prompt: 'Write summary totals',
      plan: planOf([writer]),
      context,
    });
    const violation = result.violations.find((v) => v.kind === 'uncreated-target-sheet');
    expect(violation?.detail).toContain('Summary');
    expect(violation?.fatal).toBe(false);
  });

  it('accepts a target sheet that already exists in the workbook', () => {
    const writer: SubTask = {
      id: 's5', description: 'Write totals', targetSheet: 'Sheet1',
      dependsOn: [], estimatedActions: 3,
    };
    const result = checkPlanIntegrity({ prompt: 'Write totals', plan: planOf([writer]), context });
    expect(result.violations.filter((v) => v.kind === 'uncreated-target-sheet')).toEqual([]);
  });

  it('flags an empty plan as fatal and stops there', () => {
    const result = checkPlanIntegrity({ prompt: PROMPT, plan: planOf([]), context });
    expect(result.violations).toHaveLength(1);
    expect(result.violations[0]).toMatchObject({ kind: 'empty-plan', fatal: true });
  });
});
