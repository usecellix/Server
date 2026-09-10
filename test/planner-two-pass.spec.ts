import { PlannerAgent } from '../src/agents/planner.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Two-pass planning (TASKS.md #191) — the actual fix for the truncation root
 * cause behind #170/#187/#190's live incidents. A single-pass plan for a
 * large compound build (12 month sheets + a multi-section dashboard) has to
 * describe ~19 subtasks in one JSON response and repeatedly truncated even at
 * the last-resort token ceiling. Two-pass planning identifies coarse PHASES
 * first (a handful of short entries, nearly impossible to truncate), then
 * expands each phase into real subtasks in its own small, independent call —
 * so no single call ever has to describe everything at once.
 */

const REAL_INCIDENT_PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of the ' +
  'remaining sheets, in the main sheet i need to have dashboard also, my need to record payments and related ' +
  'things ,which all month sheets include Unit No, Guest, Guest name, check in, check out, Rate per night, ' +
  'total amount, source, payment status, bank account';

function sheet(name: string): WorkbookContext['sheets'][number] {
  return {
    name,
    usedRange: 'A1',
    rowCount: 0,
    columnCount: 0,
    values: [],
    formulas: [],
    numberFormats: [],
    structure: 'unknown',
    headerRowIndex: 0,
  };
}

function emptyContext(): WorkbookContext {
  return {
    activeSheetName: 'Sheet1',
    sheets: [sheet('Sheet1')],
    namedRanges: [],
    tables: [],
  };
}

function buildAgent(completeImpl: jest.Mock): PlannerAgent {
  const llm = { complete: completeImpl } as unknown as OpenRouterService;
  const config = {
    openRouterModelHigh: 'openai/gpt-5',
    openRouterModelPlanner: 'openai/gpt-5',
  } as unknown as AppConfigService;
  return new PlannerAgent(llm, config);
}

// p1/p2/p3 deliberately target three DIFFERENT sheets — these fixtures exist
// to test sequencing, namespacing, and cross-phase stitching in isolation
// from `mergeSameSheetPhases` (which has its own dedicated tests below and
// would otherwise collapse p2+p3 if they shared a targetSheet).
const coarseResponse = JSON.stringify({
  phases: [
    {
      id: 'p1',
      kind: 'Create the 12 month sheets with standard headers',
      targetSheet: 'January',
      dependsOn: [],
      repeatFor: ['January', 'February'],
    },
    {
      id: 'p2',
      kind: 'Create the Main dashboard scaffold',
      targetSheet: 'Main',
      dependsOn: [],
    },
    {
      id: 'p3',
      kind: 'Write the Lists sheet lookup tables',
      targetSheet: 'Lists',
      dependsOn: ['p1', 'p2'],
    },
  ],
  clarificationsNeeded: [],
  confidence: 'high',
  reasoning: 'Split into month sheets, dashboard scaffold, and lookup lists.',
});

function phaseResponse(subtasks: Array<{ id: string; description: string; targetSheet: string; dependsOn?: string[] }>) {
  return JSON.stringify({
    subtasks: subtasks.map((s) => ({ ...s, dependsOn: s.dependsOn ?? [], estimatedActions: 1 })),
    clarificationsNeeded: [],
    confidence: 'high',
    reasoning: 'ok',
  });
}

describe('PlannerAgent — two-pass planning', () => {
  it('routes a long, compound prompt through the coarse+expansion calls, not a single-pass one', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(coarseResponse)
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create January', targetSheet: 'January' }]))
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create Main', targetSheet: 'Main' }]))
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Write lists', targetSheet: 'Lists' }]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_1', undefined, 3);

    // 1 coarse call + 3 phase-expansion calls (one per phase) — never one
    // giant single-pass call describing all subtasks at once.
    expect(complete).toHaveBeenCalledTimes(4);
    // 3 expanded + February cloned from January (repeatFor coverage, #229).
    expect(plan.subtasks).toHaveLength(4);
  });

  // TASKS.md #193 — the user asked for small, summarized progress updates
  // during planning instead of one static "Planning your request..." message
  // sitting there for the whole (now multi-call) two-pass sequence.
  it('reports short, incremental progress via onProgress — one update per phase, not one big block', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(coarseResponse)
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create January', targetSheet: 'January' }]))
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create Main', targetSheet: 'Main' }]))
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Write lists', targetSheet: 'Lists' }]),
      );
    const agent = buildAgent(complete);
    const progressMessages: string[] = [];

    await agent.plan(
      REAL_INCIDENT_PROMPT,
      emptyContext(),
      [],
      undefined,
      'corr_1b',
      undefined,
      3,
      undefined,
      (summary) => progressMessages.push(summary),
    );

    // One short update for the coarse pass, one per phase (3), one on
    // completion — never a single long block, and never silence for the
    // whole multi-call sequence.
    expect(progressMessages.length).toBeGreaterThanOrEqual(5);
    expect(progressMessages.every((m) => m.length < 120)).toBe(true);
    expect(progressMessages.some((m) => /step 1\/3/.test(m))).toBe(true);
    expect(progressMessages.some((m) => /step 2\/3/.test(m))).toBe(true);
    expect(progressMessages.some((m) => /step 3\/3/.test(m))).toBe(true);
  });

  it('does NOT trigger two-pass for a short prompt, even a compound one', async () => {
    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({
        subtasks: [
          { id: 's1', description: 'Build the dashboard', targetSheet: 'Dashboard', dependsOn: [], estimatedActions: 2 },
        ],
        clarificationsNeeded: [],
        confidence: 'high',
        reasoning: 'ok',
      }),
    );
    const agent = buildAgent(complete);

    await agent.plan(
      'Build a dashboard with a pivot table, a chart, and a KPI summary',
      emptyContext(),
      [],
      undefined,
      'corr_2',
      undefined,
      3,
    );

    // Exactly one call — the ordinary single-pass path, unaffected.
    expect(complete).toHaveBeenCalledTimes(1);
  });

  it('namespaces each phase\'s subtask ids so identical local ids (s1, s2) never collide across phases', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(coarseResponse)
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create January', targetSheet: 'January' }]))
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create Main', targetSheet: 'Main' }]))
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Write lists', targetSheet: 'Lists' }]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_3', undefined, 3);

    const ids = plan.subtasks.map((s) => s.id);
    expect(new Set(ids).size).toBe(ids.length); // all unique
    expect(ids).toEqual(['p1_s1', 'p1_s1_r2', 'p2_s1', 'p3_s1']);
  });

  it('wires a phase\'s subtasks (with no local dependsOn) to depend on the phases it declared as dependencies', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(coarseResponse)
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create January', targetSheet: 'January' }]))
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create Main', targetSheet: 'Main' }]))
      .mockResolvedValueOnce(
        // p3 depends on p1+p2 and its own subtask has NO local dependsOn —
        // must be wired to depend on p1's and p2's real subtask ids.
        phaseResponse([{ id: 's1', description: 'Write lists', targetSheet: 'Lists' }]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_4', undefined, 3);

    const p3Subtask = plan.subtasks.find((s) => s.id === 'p3_s1')!;
    expect(p3Subtask.dependsOn.sort()).toEqual(['p1_s1', 'p1_s1_r2', 'p2_s1']);
  });

  it('expands phases in dependency order, so a later phase sees an earlier phase\'s REAL subtask ids', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(coarseResponse)
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create January', targetSheet: 'January' }]))
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create Main', targetSheet: 'Main' }]))
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Write lists', targetSheet: 'Lists' }]),
      );
    const agent = buildAgent(complete);

    await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_5', undefined, 3);

    // Call 4 is p3's expansion (depends on p1, p2) — its user message must
    // reference the REAL namespaced subtask ids from p1/p2, not the phase ids.
    const p3Call = complete.mock.calls[3][0] as { userMessage: string };
    expect(p3Call.userMessage).toContain('p1_s1');
    expect(p3Call.userMessage).toContain('p2_s1');
  });

  it('retries a truncated phase expansion once at a larger budget, scoped to that phase only', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(coarseResponse)
      .mockResolvedValueOnce('{"subtasks": [') // p1: truncated / unparseable
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create January (retry)', targetSheet: 'January' }]))
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create Main', targetSheet: 'Main' }]))
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Write lists', targetSheet: 'Lists' }]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_6', undefined, 3);

    // 1 coarse + (p1 fail + p1 retry) + p2 + p3 = 5 calls total.
    expect(complete).toHaveBeenCalledTimes(5);
    expect(plan.subtasks.some((s) => s.description === 'Create January (retry)')).toBe(true);
  });

  it('falls back to a single all-covering phase when the coarse pass itself fails entirely', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce('not json at all {{{') // coarse attempt 1
      .mockResolvedValueOnce('still not json {{{') // coarse retry
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Do the whole thing', targetSheet: 'Sheet1' }]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_7', undefined, 3);

    expect(plan.subtasks).toHaveLength(1);
    expect(plan.subtasks[0].description).toBe('Do the whole thing');
  });

  it('prunes a phase whose expansion depends on a phase that failed entirely, same as single-pass pruning', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(coarseResponse)
      .mockResolvedValueOnce('not json {{{') // p1 attempt 1
      .mockResolvedValueOnce('still not json {{{') // p1 retry — p1 fails entirely
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create Main', targetSheet: 'Main' }]))
      .mockResolvedValueOnce(
        // p3 depends on p1 (failed) + p2 (ok) — its subtask has no local deps,
        // so it gets wired to depend on p1's (now-empty) + p2's subtask ids.
        // Since p1 produced nothing, p3's subtask ends up depending only on a
        // phase that produced no subtasks at all for THAT dependency slot —
        // pruneUnsatisfiableSubtasks must not choke on this, and p2's real
        // dependency must still be honored.
        phaseResponse([{ id: 's1', description: 'Write lists', targetSheet: 'Lists' }]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_8', undefined, 3);

    // p1 failed and produced no subtasks; p2 and p3 (depending only on p2's
    // real id, since p1 contributed nothing to stitch) still deliver.
    const ids = plan.subtasks.map((s) => s.id);
    expect(ids).toContain('p2_s1');
    expect(ids).toContain('p3_s1');
    expect(plan.clarificationsNeeded.some((c) => /could not plan/i.test(c))).toBe(true);
  });
});

describe('PlannerAgent — two-pass planning — coverage safety nets (TASKS.md #229, #230)', () => {
  const MONTHS = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
  ];

  // Replays the live 2026-09-10 run: the coarse pass planned all 12 months,
  // the expansion returned January only, Main's formulas referenced all 12,
  // and January's dropdowns pointed at a Lists sheet no phase created. The
  // workbook ended up with Main + January and 36 #REF!/#VALUE! cells.
  it('delivers all 12 month sheets and a Lists sheet when the expansion returned January only', async () => {
    const coarse = JSON.stringify({
      phases: [
        { id: 'p1', kind: 'Create the 12 month sheets', targetSheet: 'January', dependsOn: [], repeatFor: MONTHS },
        { id: 'p2', kind: 'Build the Main sheet with dashboard', targetSheet: 'Main', dependsOn: ['p1'] },
      ],
      clarificationsNeeded: [],
      confidence: 'high',
      reasoning: 'Month sheets, then Main.',
    });
    const complete = jest
      .fn()
      .mockResolvedValueOnce(coarse)
      .mockResolvedValueOnce(
        phaseResponse([
          {
            id: 's1',
            targetSheet: 'January',
            description:
              "Create sheet 'January' with headers in A1:J1, create table tblJanuary over A1:J2, add dropdowns " +
              'for Source (Lists!$B$3:$B$20), Payment Status (Lists!$C$3:$C$20)',
          },
        ]),
      )
      .mockResolvedValueOnce(
        phaseResponse([
          {
            id: 's1',
            targetSheet: 'Main',
            description: `Monthly totals: ${MONTHS.map((m, i) => `B${i + 6} =SUM(${m}!G:G)`).join(', ')}`,
          },
        ]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_12', undefined, 3);

    const sheets = plan.subtasks.map((s) => s.targetSheet);
    for (const month of MONTHS) expect(sheets).toContain(month);
    expect(sheets).toContain('Lists');

    // Main waits for every month, not just January — and its expansion call
    // was told about all 12 month subtask ids.
    const main = plan.subtasks.find((s) => s.targetSheet === 'Main')!;
    expect(main.dependsOn).toHaveLength(12);
    const mainCall = complete.mock.calls[2][0] as { userMessage: string };
    expect(mainCall.userMessage).toContain('p1_s1_r12');

    // Every month's dropdowns wait for the Lists sheet.
    const listsId = plan.subtasks.find((s) => s.targetSheet === 'Lists')!.id;
    for (const s of plan.subtasks.filter((t) => MONTHS.includes(t.targetSheet))) {
      expect(s.dependsOn).toContain(listsId);
    }
    expect(plan.clarificationsNeeded).toEqual([]);
  });
});

describe('PlannerAgent — two-pass planning — same-sheet phase merge', () => {
  // Reproduces the exact production incident: the coarse pass split ONE
  // Main-sheet build into "consolidate details into Main" (p2) and "build
  // dashboard/KPIs on Main" (p3) as separate phases. Each phase's expansion
  // call can't see what the other phase's expansion actually wrote, so both
  // independently rebuilt an identical totals table and both tried to create
  // the Main sheet — the second phase's writes were rejected by the
  // overwrite guard as duplicates. `mergeSameSheetPhases` must collapse both
  // into ONE phase before expansion ever runs, so only one expansion call
  // ever plans Main's build.
  const duplicateMainPhaseCoarseResponse = JSON.stringify({
    phases: [
      {
        id: 'p1',
        kind: 'Create the 12 monthly sheets with standard headers',
        targetSheet: 'January',
        dependsOn: [],
        repeatFor: ['January', 'February'],
      },
      {
        id: 'p2',
        kind: 'Create the Main sheet that consolidates all details from the 12 month sheets',
        targetSheet: 'Main',
        dependsOn: ['p1'],
      },
      {
        id: 'p3',
        kind: 'Build the dashboard section on the Main sheet with KPIs and breakdowns',
        targetSheet: 'Main',
        dependsOn: ['p2'],
      },
    ],
    clarificationsNeeded: [],
    confidence: 'high',
    reasoning: 'Split into the three natural units the user named.',
  });

  it('merges two non-repeatFor phases that target the same sheet into one, before any expansion runs', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(duplicateMainPhaseCoarseResponse)
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create January', targetSheet: 'January' }]))
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Build Main sheet (consolidation + dashboard)', targetSheet: 'Main' }]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_9', undefined, 3);

    // 1 coarse call + 2 expansion calls (p1, merged-p2) — NOT 3. Only one
    // call ever expands Main's build, so it can't duplicate itself.
    expect(complete).toHaveBeenCalledTimes(3);
    const ids = plan.subtasks.map((s) => s.id);
    expect(ids).toEqual(['p1_s1', 'p1_s1_r2', 'p2_s1']);
  });

  it('does not merge repeatFor phases even if they share a representative targetSheet', async () => {
    const repeatForCoarseResponse = JSON.stringify({
      phases: [
        {
          id: 'p1',
          kind: 'Create the 12 monthly sheets with standard headers',
          targetSheet: 'January',
          dependsOn: [],
          repeatFor: ['January', 'February'],
        },
        {
          id: 'p2',
          kind: 'Populate monthly summary formulas',
          targetSheet: 'January',
          dependsOn: ['p1'],
          repeatFor: ['January', 'February'],
        },
      ],
      clarificationsNeeded: [],
      confidence: 'high',
      reasoning: 'Two repeated passes over the same month sheets.',
    });
    const complete = jest
      .fn()
      .mockResolvedValueOnce(repeatForCoarseResponse)
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create January', targetSheet: 'January' }]))
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Summary formulas', targetSheet: 'January' }]));
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_10', undefined, 3);

    // Both repeatFor phases survive as separate phases despite sharing
    // "January" as their representative targetSheet — 1 coarse + 2 expansions.
    expect(complete).toHaveBeenCalledTimes(3);
    const ids = plan.subtasks.map((s) => s.id);
    expect(ids).toEqual(['p1_s1', 'p1_s1_r2', 'p2_s1', 'p2_s1_r2']);
  });

  it('unions dependsOn and remaps references to the surviving phase id when merging', async () => {
    const threeWayCoarseResponse = JSON.stringify({
      phases: [
        { id: 'p1', kind: 'Create Lists sheet', targetSheet: 'Lists', dependsOn: [] },
        { id: 'p2', kind: 'Consolidate details into Main', targetSheet: 'Main', dependsOn: ['p1'] },
        { id: 'p3', kind: 'Build dashboard on Main', targetSheet: 'Main', dependsOn: ['p1'] },
        { id: 'p4', kind: 'Add a summary note referencing Main', targetSheet: 'Notes', dependsOn: ['p3'] },
      ],
      clarificationsNeeded: [],
      confidence: 'high',
      reasoning: 'Four units, two of which collide on Main.',
    });
    const complete = jest
      .fn()
      .mockResolvedValueOnce(threeWayCoarseResponse)
      .mockResolvedValueOnce(phaseResponse([{ id: 's1', description: 'Create Lists', targetSheet: 'Lists' }]))
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Build Main sheet', targetSheet: 'Main' }]),
      )
      .mockResolvedValueOnce(
        phaseResponse([{ id: 's1', description: 'Add summary note', targetSheet: 'Notes' }]),
      );
    const agent = buildAgent(complete);

    const plan = await agent.plan(REAL_INCIDENT_PROMPT, emptyContext(), [], undefined, 'corr_11', undefined, 3);

    // p2+p3 merge into "p2" (the survivor, first in the group). p4 depended
    // on p3 — that reference must be remapped to the surviving p2 id so p4's
    // subtask still ends up depending on Main's real subtask ids.
    expect(complete).toHaveBeenCalledTimes(4);
    const notesSubtask = plan.subtasks.find((s) => s.id === 'p4_s1')!;
    expect(notesSubtask.dependsOn).toEqual(['p2_s1']);
  });
});
