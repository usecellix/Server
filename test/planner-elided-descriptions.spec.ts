import { PlannerAgent } from '../src/agents/planner.agent';
import { PlannerOutput, SubTask } from '../src/agents/types/agent.types';

/**
 * TASKS.md #271 — the live failure: a 95-subtask plan (each month split seven
 * ways) blew the completion budget, and the model shortened every description
 * instead of emitting fewer subtasks. "Write headers... ", "Set formulas... ",
 * median 20 characters. Those parse fine and carry real ids and dependsOn, so
 * every existing guard passed them — and the Executor, given no schema,
 * invented one per sheet (GST purchase registers, a four-column expense log,
 * and five sheets with no header row at all).
 */
describe('PlannerAgent — elided descriptions guard (TASKS.md #271)', () => {
  const agent = new PlannerAgent({} as never, {} as never);

  const assertNotElided = (plan: PlannerOutput) =>
    (agent as unknown as {
      assertDescriptionsNotElided: (o: PlannerOutput, m: string) => void;
    }).assertDescriptionsNotElided(plan, 'build me a ledger');

  const subtask = (id: string, description: string, targetSheet = 'January'): SubTask => ({
    id,
    description,
    targetSheet,
    dependsOn: [],
    estimatedActions: 5,
  });

  const plan = (subtasks: SubTask[]): PlannerOutput => ({
    subtasks,
    clarificationsNeeded: [],
    confidence: 'high',
    reasoning: '',
  });

  /** The real stub set, verbatim from the failing run's plan. */
  const livePlan = () =>
    plan([
      subtask('p2_s1', "Create sheet 'January'... "),
      subtask('p2_s2', 'Write headers... '),
      subtask('p2_s3', 'Create table tblJan... '),
      subtask('p2_s4', 'Set formulas... '),
      subtask('p2_s5', 'Add data validation dropdowns... '),
      subtask('p2_s6', 'Set column widths... '),
      subtask('p2_s7', 'Apply font...'),
    ]);

  it('refuses the live elided plan rather than letting the Executor invent schemas', () => {
    expect(() => assertNotElided(livePlan())).toThrow(/elided/i);
  });

  it('names how many were elided, so the log says what actually happened', () => {
    expect(() => assertNotElided(livePlan())).toThrow(/7 elided subtask descriptions out of 7/);
  });

  it('accepts a real plan whose descriptions actually carry the schema', () => {
    expect(() =>
      assertNotElided(
        plan([
          subtask(
            'p2_s1',
            "Create sheet 'January', write headers A1:M1 [Unit No, Guest, Guest Name, Check In, " +
              'Check Out, Nights, Rate Per Night, Total Amount, Source, Payment Status, Amount ' +
              'Received, Balance Due, Bank Account], create table tblJanuary over A1:M2 with ' +
              'showFilterButton false, and set F2 =IF(OR(D2="",E2=""),"",E2-D2)',
          ),
          subtask(
            'p2_s2',
            "Create sheet 'February' with the same 13-column header row and the same row-2 formulas",
            'February',
          ),
        ]),
      ),
    ).not.toThrow();
  });

  /**
   * "Hide the Lists sheet" is 20 characters AND a complete instruction. The
   * signal is a plan-WIDE pattern, not any single terse subtask — flagging
   * those would refuse perfectly good plans.
   */
  it('tolerates a few genuinely terse but complete subtasks among real ones', () => {
    const long = "Create sheet 'January' and write the full 13-column header row A1:M1 with row-2 formulas";
    expect(() =>
      assertNotElided(
        plan([
          subtask('s1', long),
          subtask('s2', long, 'February'),
          subtask('s3', long, 'March'),
          subtask('s4', long, 'April'),
          subtask('s5', 'Hide the Lists sheet', 'Lists'),
        ]),
      ),
    ).not.toThrow();
  });

  it('does not fire on a small plan of two short subtasks (below the count floor)', () => {
    expect(() =>
      assertNotElided(plan([subtask('s1', 'Sort by date'), subtask('s2', 'Freeze row 1')])),
    ).not.toThrow();
  });

  it('treats an empty description as elided', () => {
    const long = "Create sheet 'January' and write the full 13-column header row A1:M1 with row-2 formulas";
    expect(() =>
      assertNotElided(
        plan([subtask('s1', ''), subtask('s2', '   '), subtask('s3', ''), subtask('s4', long)]),
      ),
    ).toThrow(/elided/i);
  });
});
