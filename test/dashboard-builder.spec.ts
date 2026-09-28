import { splitSpecPinnedSubtasks, buildHeaderTableActions } from '../src/agents/utils/header-table-split.util';
import { applyConsolidationPass } from '../src/excel-ai/utils/consolidation-pass.util';
import { applyBuildSpecToSubtasks } from '../src/agents/utils/build-spec.util';
import { checkPlanIntegrity } from '../src/agents/utils/plan-integrity.util';
import { normalizeExecutorOutput } from '../src/agents/utils/normalize-executor-output.util';
import { reconcileRun } from '../src/agents/utils/reconcile.util';
import { replaceDashboardSubtasks } from '../src/agents/utils/dashboard-builder.util';
import { isDeterministicStep, PlannerOutput, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #327 — the ledger's Main dashboard, built by code.
 *
 * Live run_1790331172604_c0njd8s reported "All steps applied" over a Main that
 * had: July–December never written (`operations: [24]`), a KPI band summing
 * B5:B10 when told B5:B16, no title, a Pending column that SUMIF'd a status
 * the Lists sheet never offers, and a chart plotting "Month" as a series.
 * Every one of those is a model slip on work that is fully determined once the
 * month sheets' header row is known.
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

const month = (name: string, i: number): SubTask => ({
  id: `p2_s${i + 1}`,
  targetSheet: name,
  dependsOn: ['p1_s1'],
  estimatedActions: 26,
  description:
    `Create sheet '${name}'. Write headers in row 1 (A1:M1): Unit No, Guest, Guest Name, Check In, ` +
    `Check Out, Nights, Rate Per Night, Total Amount, Source, Payment Status, Amount Received, ` +
    `Balance Due, Bank Account. Create Excel Table 'tbl${name}' over A1:M2.`,
});
const main = (id: string, description: string, dependsOn: string[] = []): SubTask => ({
  id, targetSheet: 'Main', description, dependsOn, estimatedActions: 6,
});

// The live plan's Main subtasks, near-verbatim.
const livePlan: PlannerOutput = {
  subtasks: [
    { id: 'p1_s1', description: "Create sheet 'Lists'", targetSheet: 'Lists', dependsOn: [], estimatedActions: 4 },
    ...MONTHS.map(month),
    main('p3_s1', "Create sheet 'Main' at position 0 with title 'Payment Dashboard' in A1", ['p2_s1']),
    main('p3_s2', 'Monthly Totals A4:D16, January–June: B5 =SUM(January!H:H) ...', ['p3_s1']),
    main('p3_s3', 'Monthly Totals rows 11–16, July–December ...', ['p3_s2']),
    main('p3_s4', 'KPI band rows 1–2: B2 =SUM(B5:B16) ...', ['p3_s3']),
    main('p3_s6', 'Chart of Main!A4:D16 at F4', ['p3_s3']),
    main('p3_s7', 'Hide gridlines on Main and set column widths', ['p3_s4', 'p3_s6']),
  ],
  clarificationsNeeded: [],
  confidence: 'high',
  reasoning: '',
};

function pipeline(plan: PlannerOutput = livePlan) {
  const gated = checkPlanIntegrity({ prompt: PROMPT, plan, context }).plan;
  const pinned = applyBuildSpecToSubtasks(gated.subtasks, { sheets: [{ names: MONTHS, columns: USER_COLUMNS }] });
  return replaceDashboardSubtasks(splitSpecPinnedSubtasks(pinned));
}

type Op = { address: string; value?: string; formula?: string };
function cells(subtasks: SubTask[]): Map<string, Op> {
  const dash = subtasks.find((s) => s.deterministicActions?.length)!;
  const batch = dash.deterministicActions!.find((a) => a.type === 'BATCH_SET') as unknown as { operations: Op[] };
  return new Map(batch.operations.map((op) => [op.address, op]));
}

describe('ledger dashboard — built by code (TASKS.md #327)', () => {
  it('replaces every model-planned Main subtask with ONE deterministic step', () => {
    const { subtasks, replaced } = pipeline();
    expect(replaced.sort()).toEqual(['p3_s1', 'p3_s2', 'p3_s3', 'p3_s4', 'p3_s6', 'p3_s7']);
    const onMain = subtasks.filter((s) => s.targetSheet === 'Main');
    expect(onMain).toHaveLength(1);
    expect(isDeterministicStep(onMain[0])).toBe(true);
    // It waits for the month sheets to EXIST (their header steps), not for their formulas (#309).
    expect(onMain[0].dependsOn).toEqual(MONTHS.map((m) => `hdr_${m}`));
  });

  it('writes all twelve months — the row the live run lost (July–December) included', () => {
    const c = cells(pipeline().subtasks);
    MONTHS.forEach((m, i) => expect(c.get(`A${9 + i}`)?.value).toBe(m));
    expect(c.get('A21')?.value).toBe('Total');
  });

  it('reads Total Amount and Payment Status from their REAL columns in the built header row', () => {
    // Resolved layout: ... G Rate Per Night, H Total Amount, I Source, J Payment Status ...
    const c = cells(pipeline().subtasks);
    expect(c.get('B9')?.formula).toBe('=SUM(January!H:H)');
    expect(c.get('C9')?.formula).toBe('=SUMIF(January!J:J,"Paid",January!H:H)');
    expect(c.get('E20')?.formula).toBe('=COUNT(December!H:H)');
  });

  it('Pending is what is not yet Paid — never a SUMIF on a status word the list may not contain', () => {
    const c = cells(pipeline().subtasks);
    expect(c.get('D9')?.formula).toBe('=B9-C9');
    const all = [...c.values()].map((op) => op.formula ?? '').join(' ');
    expect(all).not.toMatch(/"Pending"/);
  });

  it('the Total row sums all twelve months and the KPI band reads the Total row', () => {
    const c = cells(pipeline().subtasks);
    expect(c.get('B21')?.formula).toBe('=SUM(B9:B20)');
    expect(c.get('B5')?.formula).toBe('=B21');
    expect(c.get('E5')?.formula).toBe('=E21');
    expect(c.get('A1')?.value).toBe('Payments Dashboard');
  });

  it('charts the months only — full month names as categories, never the Total row', () => {
    const dash = pipeline().subtasks.find((s) => s.deterministicActions?.length)!;
    const chart = dash.deterministicActions!.find((a) => a.type === 'CREATE_CHART') as unknown as Record<string, string>;
    expect(chart.sourceRange).toBe('A8:D20');
    expect(chart.startCell).toBe('G7');
  });

  it('every action survives normalization — the client receives exactly what was built', () => {
    const dash = pipeline().subtasks.find((s) => s.deterministicActions?.length)!;
    const out = normalizeExecutorOutput({ subtaskId: dash.id, actions: dash.deterministicActions! }, dash);
    expect(out.droppedActions).toEqual([]);
    expect(out.actions).toHaveLength(dash.deterministicActions!.length);
    expect((out.actions[0] as { name?: string }).name).toBe('Main');
  });

  it('a full build of headers + dashboard reconciles to no gaps and no dangling references', () => {
    const { subtasks } = pipeline();
    const applied = subtasks.flatMap((s) =>
      s.isDeterministicHeaderStep ? buildHeaderTableActions(s) : s.deterministicActions ?? [],
    );
    const { gaps } = reconcileRun({ subtasks, appliedActions: applied, preExistingSheets: ['Sheet1', 'Lists'] });
    expect(gaps.filter((g) => g.kind === 'dangling-reference' || g.kind === 'missing-sheet')).toEqual([]);
  });

  // TASKS.md #333 — "all the details of the remaining sheets": Main also lists
  // every booking, live. The user's words after seeing a Main without it:
  // "it only updated in the chart and the amount section, below it is blank".
  it('adds an All Bookings section headed by Month + the month sheets\' exact columns', () => {
    const c = cells(pipeline().subtasks);
    expect(c.get('A23')?.value).toBe('All Bookings');
    expect(c.get('A24')?.value).toBe('Month');
    expect(c.get('B24')?.value).toBe('Unit No');
    expect(c.get('N24')?.value).toBe('Bank Account'); // 13 month columns after Month
    // Nothing is written into the rows the list will fill.
    expect([...c.keys()].filter((address) => /^[A-Z]+(2[5-9]|[3-9]\d)$/.test(address))).toEqual([]);
  });

  it('the consolidation pass fills that section with ONE live formula at A25', () => {
    const { subtasks } = pipeline();
    const applied = subtasks.flatMap((s) =>
      s.isDeterministicHeaderStep ? buildHeaderTableActions(s) : s.deterministicActions ?? [],
    );
    const added = applyConsolidationPass(applied, { dynamicArrays: true }).slice(applied.length) as unknown as Array<
      Record<string, unknown>
    >;
    expect(added).toHaveLength(1);
    expect(added[0]).toMatchObject({ type: 'SET_FORMULA', sheetName: 'Main', row: 24, col: 0 });
    expect(added[0].formula).toContain('January!A2:M500');
    expect(added[0].formula).toContain('December!A2:M500');
  });

  it('keeps the chart above the All Bookings section', () => {
    const dash = pipeline().subtasks.find((s) => s.deterministicActions?.length)!;
    const chart = dash.deterministicActions!.find((a) => a.type === 'CREATE_CHART') as unknown as Record<string, string>;
    expect(chart.endCell).toMatch(/22$/);
  });

  it('leaves a plan alone when there is no dashboard sheet or no Total column', () => {
    const noMain: PlannerOutput = { ...livePlan, subtasks: livePlan.subtasks.filter((s) => s.targetSheet !== 'Main') };
    expect(pipeline(noMain).shape).toBeNull();

    const pinned = applyBuildSpecToSubtasks(
      [{ id: 'x', targetSheet: 'Main', description: 'dashboard', dependsOn: [], estimatedActions: 1 },
        ...['A', 'B', 'C'].map((n, i) => ({ ...month(n, i), description: `Create sheet '${n}' with headers Name, Qty` }))],
      { sheets: [{ names: ['A', 'B', 'C'], columns: ['Name', 'Qty'] }] },
    );
    expect(replaceDashboardSubtasks(splitSpecPinnedSubtasks(pinned)).shape).toBeNull();
  });
});
