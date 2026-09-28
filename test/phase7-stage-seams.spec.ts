import {
  buildHeaderTableActions,
  resolveHeaderRow,
  splitSpecPinnedSubtasks,
  stripAlreadyBuiltInstructions,
} from '../src/agents/utils/header-table-split.util';
import { applyBuildSpecToSubtasks } from '../src/agents/utils/build-spec.util';
import { checkPlanIntegrity } from '../src/agents/utils/plan-integrity.util';
import { normalizeExecutorOutput } from '../src/agents/utils/normalize-executor-output.util';
import { reconcileRun } from '../src/agents/utils/reconcile.util';
import { Action, PlannerOutput, SubTask, WorkbookContext } from '../src/agents/types/agent.types';
import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 7 — the seams between stages.
 *
 * This suite exists because of the pattern in the plan's §0: six consecutive
 * fixes in one session each solved one stage and created a precondition the
 * NEXT stage could not satisfy, and nothing caught any of them except the next
 * live run.
 *
 *   #283  header row narrower than the rest step's own instruction
 *   #284  parser handled one phrasing; the next run used another
 *   #289  table spanned A1:J2 but row 2 had no cells, so every formula was refused
 *   #290  a formula fragment became a real sheet name
 *   #279  server emitted ADD_SHEET with `sheetName`; the client reads only `name`
 *
 * Each assertion below is a POSTCONDITION one stage owes the next. They are
 * grouped by seam and named for what breaks if they fail, so a future change
 * that re-opens one of these is told immediately instead of in a workbook.
 */

const USER_COLUMNS = [
  'Unit No', 'Guest', 'Guest name', 'check in', 'check out',
  'Rate per night', 'total amount', 'source', 'payment status', 'bank account',
];

const PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of ' +
  'the remaining sheets, in the main sheet i need to have dashboard also, my need to record payments and ' +
  'related things ,which all month sheets include ' + USER_COLUMNS.join(', ');

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

/** A realistic planner subtask, in the phrasing a live run actually produced. */
const plannedMonth = (month: string, i: number): SubTask => ({
  id: `p2_s${i + 1}`,
  targetSheet: month,
  dependsOn: ['p1_s1'],
  estimatedActions: 26,
  description:
    `Create sheet '${month}'. Write headers in row 1 (A1:M1): Unit No, Guest, Guest Name, Check In, ` +
    `Check Out, Nights, Rate Per Night, Total Amount, Source, Payment Status, Amount Received, ` +
    `Balance Due, Bank Account. Create Excel Table 'tbl${month}' over A1:M2. Set row-2 formulas: ` +
    `F2 =IF(OR(D2="",E2=""),"",E2-D2). Set column widths: A=10, B=12.`,
});

const basePlan: PlannerOutput = {
  subtasks: [
    { id: 'p1_s1', description: "Create sheet 'Lists'", targetSheet: 'Lists', dependsOn: [], estimatedActions: 4 },
    ...MONTHS.map(plannedMonth),
  ],
  clarificationsNeeded: [],
  confidence: 'high',
  reasoning: '',
};

const spec = { sheets: [{ names: MONTHS, columns: USER_COLUMNS }] };

/** The full pipeline a plan travels, in order. */
function runPipeline() {
  const gated = checkPlanIntegrity({ prompt: PROMPT, plan: basePlan, context }).plan;
  const pinned = applyBuildSpecToSubtasks(gated.subtasks, spec);
  const split = splitSpecPinnedSubtasks(pinned);
  return { gated, pinned, split };
}

describe('Phase 7 seam: planner -> integrity gate', () => {
  it('a complete plan passes through unchanged', () => {
    const { gated } = runPipeline();
    expect(gated.subtasks).toHaveLength(13);
  });
});

describe('Phase 7 seam: spec pinning -> deterministic split', () => {
  it('every month subtask carries the user’s exact columns after pinning', () => {
    const { pinned } = runPipeline();
    const months = pinned.filter((s) => MONTHS.includes(s.targetSheet));
    expect(months).toHaveLength(12);
    expect(months.every((s) => s.expectedHeaders?.length === USER_COLUMNS.length)).toBe(true);
  });

  it('the split produces one deterministic step per month plus its lighter rest step', () => {
    const { split } = runPipeline();
    expect(split.filter((s) => s.isDeterministicHeaderStep)).toHaveLength(12);
    expect(split.filter((s) => MONTHS.includes(s.targetSheet) && !s.isDeterministicHeaderStep)).toHaveLength(12);
  });

  it('the rest step keeps its ORIGINAL id, so anything depending on it still resolves', () => {
    const { split } = runPipeline();
    const ids = new Set(split.map((s) => s.id));
    for (let i = 0; i < 12; i += 1) expect(ids.has(`p2_s${i + 1}`)).toBe(true);
  });
});

describe('Phase 7 seam: deterministic step -> the formula step that follows it (#283, #289)', () => {
  it('builds the FULL planned layout, not just the user’s columns (#283)', () => {
    const { split } = runPipeline();
    const header = split.find((s) => s.isDeterministicHeaderStep && s.targetSheet === 'January')!;
    // The planner's own description names 13 columns; the user named 10.
    expect(resolveHeaderRow(header)).toHaveLength(13);
  });

  it('SEEDS THE DATA ROW — a table whose row 2 has no cells blocks every formula (#289)', () => {
    const { split } = runPipeline();
    const header = split.find((s) => s.isDeterministicHeaderStep && s.targetSheet === 'January')!;
    const actions = buildHeaderTableActions(header) as Array<Record<string, any>>;
    const ops = actions[1].operations as Array<{ address: string }>;

    const rowTwo = ops.filter((op) => /^[A-Z]+2$/.test(op.address));
    expect(rowTwo.length).toBeGreaterThan(0);

    // The table's declared range must not extend past the rows that exist.
    const tableRange = actions[2].range as string;
    const lastRow = Number(/(\d+)$/.exec(tableRange)![1]);
    const rowsWritten = new Set(ops.map((op) => /(\d+)$/.exec(op.address)![1])).size;
    expect(rowsWritten).toBeGreaterThanOrEqual(lastRow);
  });

  it('the rest step is never told to rebuild what the deterministic step just built (#284)', () => {
    const { split } = runPipeline();
    const rest = split.find((s) => s.targetSheet === 'January' && !s.isDeterministicHeaderStep)!;
    expect(rest.description).not.toMatch(/Write headers in row 1/i);
    expect(rest.description).toMatch(/do NOT use ADD_SHEET, CREATE_TABLE or INSERT_COLUMN/);
    // ...but the real remaining work survives.
    expect(rest.description).toMatch(/Set row-2 formulas/i);
  });

  it('stripping never removes a sentence that also carries real work', () => {
    const mixed = "Create sheet 'January' and set column widths A=10 and apply font Aptos Narrow.";
    expect(stripAlreadyBuiltInstructions(mixed)).toBe(mixed);
  });
});

describe('Phase 7 seam: emitted action -> client handler contract (#279)', () => {
  const subtask: SubTask = {
    id: 's1', description: 'x', targetSheet: 'January', dependsOn: [], estimatedActions: 1,
  };

  it('ADD_SHEET always carries `name`, which is the only field the client handler reads', () => {
    // The model legitimately emits either shape; normalization owes the client `name`.
    for (const emitted of [
      { type: 'ADD_SHEET', sheetName: 'January' },
      { type: 'ADD_SHEET', name: 'January' },
      { type: 'CREATE_SHEET', sheetName: 'January' },
    ]) {
      const result = normalizeExecutorOutput({ subtaskId: 's1', actions: [emitted] }, subtask);
      expect((result.actions[0] as { name?: string }).name).toBe('January');
    }
  });

  it('the deterministic builder also satisfies that contract', () => {
    const header: SubTask = { ...subtask, expectedHeaders: USER_COLUMNS };
    const [addSheet] = buildHeaderTableActions(header) as Array<Record<string, any>>;
    expect(addSheet.type).toBe('ADD_SHEET');
    expect(addSheet.name).toBe('January');
  });

  it('CREATE_TABLE carries the fields its handler requires', () => {
    const header: SubTask = { ...subtask, expectedHeaders: USER_COLUMNS };
    const table = (buildHeaderTableActions(header) as Array<Record<string, any>>)[2];
    expect(table.tableName).toBeTruthy();
    expect(table.range).toMatch(/^A1:[A-Z]+\d+$/);
    expect(typeof table.hasHeaders).toBe('boolean');
  });
});

describe('Phase 7 seam: applied actions -> reconciliation (#290)', () => {
  it('a clean build reconciles to nothing', () => {
    const { split } = runPipeline();
    const applied = split
      .filter((s) => s.isDeterministicHeaderStep)
      .flatMap((s) => buildHeaderTableActions(s));

    const { gaps } = reconcileRun({
      subtasks: split.filter((s) => s.isDeterministicHeaderStep),
      appliedActions: applied,
    });
    expect(gaps).toEqual([]);
  });

  it('a formula fragment is never mistaken for a missing sheet', () => {
    const main: SubTask = {
      id: 'p3', description: 'Build Main', targetSheet: 'Main', dependsOn: [], estimatedActions: 4,
    };
    const applied = [
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      {
        type: 'BATCH_SET',
        sheetName: 'Main',
        operations: [
          { address: 'B2', formula: '=SUM(INDIRECT("\'"&TEXT(DATE(2026,ROW(INDIRECT("1:12")),1),"mmmm")&"\'!H:H"))' },
        ],
      },
    ] as unknown as Action[];

    const { gaps } = reconcileRun({ subtasks: [main], appliedActions: applied });
    expect(gaps.filter((g) => g.kind === 'dangling-reference')).toEqual([]);
  });
});

/**
 * The seam this suite was supposed to cover and did not.
 *
 * #296, found by the smoke run after #294 shipped: the Main subtask emits its
 * OWN create and its own formulas in one batch, and pre-apply formula
 * validation refused every formula with `points to unknown sheet "Main"` —
 * `virtualApply` created Main in the shadow, `shadowAsContext` dropped it
 * again. Two retries, subtask failed, dependents gated off, run ended at wave
 * 4 of 6 with Main never built.
 *
 * The seam is executor output -> formula pre-validation, and nothing here
 * crossed it. Phase 7 asserted what the SPLIT owes the formula step; it never
 * asserted that an emitted batch survives the validator standing between them.
 */
describe('Phase 7 seam: emitted actions -> formula pre-validation (#296)', () => {
  const validator = new FormulaValidatorService();

  /** The live Main subtask, reduced to the shape that failed. */
  const mainActions: Action[] = [
    { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
    { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'Payments Dashboard' },
    { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(B5:B16)' },
    { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 3, formula: '=SUM(C5:C16)' },
    { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 5, formula: '=SUM(D5:D16)' },
  ];

  it('a subtask may reference the sheet it is itself creating (#296)', () => {
    const result = validator.validatePreApply(
      mainActions,
      context,
      'Main',
      buildShadowWorkbook(context),
    );
    expect(
      result.issues.filter((i) => i.severity === 'error').map((i) => i.message),
    ).toEqual([]);
    expect(result.passed).toBe(true);
  });

  it('a consolidation formula may reference month sheets the same run creates', () => {
    // What the dashboard is ultimately for: Main reading the month sheets that
    // the deterministic header steps built one wave earlier.
    const { split } = runPipeline();
    const january = split.find((s) => s.isDeterministicHeaderStep && s.targetSheet === 'January');
    expect(january).toBeDefined();

    const built = buildHeaderTableActions(january as SubTask);
    const consolidation: Action[] = [
      ...built,
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      { type: 'SET_FORMULA', sheetName: 'Main', row: 4, col: 1, formula: "=SUM(January!H2:H100)" },
    ];

    const result = validator.validatePreApply(
      consolidation,
      context,
      'Main',
      buildShadowWorkbook(context),
    );
    expect(
      result.issues.some((i) => i.severity === 'error' && i.message.includes('unknown sheet')),
    ).toBe(false);
  });

  it('a reference to a sheet NOTHING in the run touches is still an error', () => {
    // The relaxation must not hollow out the check it relaxes.
    const result = validator.validatePreApply(
      [
        { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: "=SUM(Quarter4!B5:B16)" },
      ],
      context,
      'Main',
      buildShadowWorkbook(context),
    );
    expect(result.passed).toBe(false);
    expect(
      result.issues.some((i) => i.message.includes('unknown sheet "Quarter4"')),
    ).toBe(true);
  });

  it('bounds on a sheet that already exists are still a hard error', () => {
    // Sheet1 is real and one row tall; nothing in this batch is building it,
    // so a reference past its end is the genuine mistake the check exists for.
    const result = validator.validatePreApply(
      [{ type: 'SET_FORMULA', sheetName: 'Sheet1', row: 0, col: 1, formula: '=SUM(A50:A99)' }],
      context,
      'Sheet1',
      buildShadowWorkbook(context),
    );
    expect(result.passed).toBe(false);
  });
});
