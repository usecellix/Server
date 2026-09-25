import { ConversationService } from '../src/excel-ai/services/conversation.service';
import { Tier0DirectService } from '../src/excel-ai/services/tier0-direct.service';
import { FastifyReply } from 'fastify';
import * as sseUtil from '../src/excel-ai/utils/sse.util';
import {
  applyConsolidationPass,
  consolidationSpillRegion,
  stripSpillCollisions,
} from '../src/excel-ai/utils/consolidation-pass.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #311 / #312 — live run run_1790318441177_qx7ujmb (Sept 25, 2026).
 *
 * Main's header step, told "header row only", also wrote Nights / Total Amount
 * / Balance Due formulas into G19, I19, M19. The consolidation pass wrote its
 * spilling formula at A19 in the same wave, so A19 showed `#SPILL!` (#311).
 *
 * The formula was then re-emitted on step 8 of 8: it had been pushed into the
 * wave's flat action list but never into any subtask's recorded actions, and
 * the "already written" check reads recorded actions. The client's overwrite
 * guard refused the rewrite onto `#SPILL!` on every Accept, and the run was
 * stranded (#312).
 */

const MONTHS = ['January', 'February'];
const SCHEMA = ['Unit No', 'Guest', 'Nights', 'Total Amount'];

function monthActions(month: string): SheetActionPayload[] {
  return [
    { type: 'ADD_SHEET', name: month },
    {
      type: 'BATCH_SET',
      sheetName: month,
      operations: SCHEMA.map((label, i) => ({ address: `${String.fromCharCode(65 + i)}1`, value: label })),
    },
  ] as SheetActionPayload[];
}

/** Main's consolidated header at row 18 plus the live run's stray row-19 formulas. */
function mainHeaderWave(): SheetActionPayload[] {
  return [
    { type: 'ADD_SHEET', name: 'Main' },
    {
      type: 'BATCH_SET',
      sheetName: 'Main',
      operations: ['Month', ...SCHEMA].map((label, i) => ({
        address: `${String.fromCharCode(65 + i)}18`,
        value: label,
      })),
    },
    { type: 'SET_FORMULA', sheetName: 'Main', address: 'D19', row: 18, col: 3, formula: '=IF(B19="","",C19)' },
    { type: 'SET_FORMULA', sheetName: 'Main', address: 'E19', row: 18, col: 4, formula: '=D19*2' },
  ] as SheetActionPayload[];
}

describe('consolidation spill collisions — util (TASKS.md #311)', () => {
  const batch = (): SheetActionPayload[] => [...MONTHS.flatMap(monthActions), ...mainHeaderWave()];

  it('the spill region covers the origin column plus the source schema, from the formula row down', () => {
    expect(consolidationSpillRegion(batch())).toEqual({ sheetName: 'Main', row: 18, col: 0, lastCol: 4 });
  });

  it('removes value writes inside the region and keeps the header row, other sheets and cells outside it', () => {
    const region = consolidationSpillRegion(batch())!;
    const extra = [
      { type: 'SET_CELL', sheetName: 'Main', row: 18, col: 6, value: 'beside the spill' },
      { type: 'FORMAT_RANGE', sheetName: 'Main', range: 'A19:E40' },
      { type: 'SET_CELL', sheetName: 'January', row: 18, col: 1, value: 'other sheet' },
      {
        type: 'BATCH_SET',
        sheetName: 'Main',
        operations: [
          { address: 'B25', value: 'inside' },
          { address: 'H25', value: 'outside' },
        ],
      },
    ] as SheetActionPayload[];

    const { kept, removed } = stripSpillCollisions([...batch(), ...extra], region);

    expect(removed).toBe(3); // D19, E19, B25
    const mainFormulas = kept.filter((a) => a.type === 'SET_FORMULA' && a.sheetName === 'Main');
    expect(mainFormulas).toHaveLength(0);
    expect(kept).toContainEqual(extra[0]);
    expect(kept).toContainEqual(extra[1]);
    expect(kept).toContainEqual(extra[2]);
    expect(kept).toContainEqual({ ...extra[3], operations: [{ address: 'H25', value: 'outside' }] });
    // The header row is row 18 — above the region — and survives whole.
    const header = kept.find(
      (a) => a.type === 'BATCH_SET' && a.sheetName === 'Main',
    ) as { operations: unknown[] };
    expect(header.operations).toHaveLength(SCHEMA.length + 1);
  });

  it('the one-shot pass clears the spill area as it adds the formula', () => {
    const out = applyConsolidationPass(batch(), { dynamicArrays: true });
    const onMain = out.filter((a) => a.type === 'SET_FORMULA' && a.sheetName === 'Main');
    expect(onMain).toHaveLength(1);
    expect(onMain[0]).toMatchObject({ row: 18, col: 0 });
    expect(String(onMain[0].formula)).toContain('VSTACK');
  });

  it('leaves the batch untouched on a legacy host, where no formula spills', () => {
    const out = applyConsolidationPass(batch(), { dynamicArrays: false });
    expect(out.slice(0, batch().length)).toEqual(batch());
  });
});

describe('consolidation spill collisions — stepwise path (TASKS.md #311, #312)', () => {
  let service: ConversationService;
  let agentRunState: Record<string, jest.Mock>;
  let orchestrator: { runStepwiseWave: jest.Mock };
  let changeSetService: { createPreview: jest.Mock };
  const reply = {} as FastifyReply;
  const emit = () => undefined;

  beforeEach(() => {
    agentRunState = {
      markStatus: jest.fn().mockResolvedValue(undefined),
      summarizeSkipped: jest.fn().mockReturnValue([]),
      nextExecutableWave: jest.fn(),
      recordWaveResult: jest.fn().mockResolvedValue(undefined),
    };
    orchestrator = { runStepwiseWave: jest.fn() };
    changeSetService = {
      createPreview: jest.fn().mockResolvedValue({ changeSetId: 'cs', changes: [], irreversibleActionTypes: [] }),
    };
    service = new ConversationService(
      { updateOne: jest.fn().mockResolvedValue(undefined) } as never,
      {} as never,
      {} as never,
      {} as never,
      changeSetService as never,
      {} as never,
      orchestrator as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      new Tier0DirectService(),
      {} as never,
      {} as never,
      {} as never,
      { appendNode: jest.fn(), setMeta: jest.fn() } as never,
      {} as never,
      { debit: jest.fn().mockResolvedValue({ debited: false }) } as never,
      agentRunState as never,
    );
    jest.spyOn(service as never, 'markCompleted' as never).mockResolvedValue(undefined as never);
    jest.spyOn(sseUtil, 'endSseResponse').mockImplementation(() => undefined);
    jest.spyOn(sseUtil, 'createRequestAbortSignal').mockReturnValue({ aborted: false } as AbortSignal);
  });

  afterEach(() => jest.restoreAllMocks());

  function runWith(subtaskStates: unknown[]) {
    return {
      runId: 'run_1',
      conversationId: 'conv-1',
      prompt: 'build the ledger',
      context: { activeSheetName: 'Sheet1' },
      excelCapabilities: { dynamicArrays: true },
      subtasks: [
        ...MONTHS.map((m) => ({ id: `s_${m}`, targetSheet: m, description: m, dependsOn: [] })),
        { id: 's_main', targetSheet: 'Main', description: 'header', dependsOn: [] },
        { id: 's_fmt', targetSheet: 'Main', description: 'format', dependsOn: [] },
      ],
      subtaskStates,
      waveIndex: 1,
      waveTotal: 3,
      changeSetIds: [],
      save: jest.fn().mockResolvedValue(undefined),
    };
  }

  const monthStates = () =>
    MONTHS.map((m) => ({ subtaskId: `s_${m}`, completed: true, actions: monthActions(m) }));

  function waveOf(subtaskId: string, actions: SheetActionPayload[]) {
    return {
      actions: [...actions],
      completedSubtasks: [{ subtaskId, actions: [...actions], verified: true }],
      failedSubtask: null,
      failedSubtasks: [],
      verifierPassed: true,
    };
  }

  const exec = (run: unknown) =>
    (service as any).executeStepwiseWave(run, reply, emit, {} as never, {} as never);
  const previewed = (call = 0) =>
    changeSetService.createPreview.mock.calls[call][0].actions as SheetActionPayload[];

  it('strips the header wave\'s writes inside the spill area and ships the formula', async () => {
    const run = runWith(monthStates());
    agentRunState.nextExecutableWave.mockReturnValue({ waveIndex: 1, subtasks: [run.subtasks[2]] });
    orchestrator.runStepwiseWave.mockResolvedValue(waveOf('s_main', mainHeaderWave()));

    await exec(run);

    const onMain = previewed().filter((a) => a.type === 'SET_FORMULA' && a.sheetName === 'Main');
    expect(onMain).toHaveLength(1);
    expect(onMain[0]).toMatchObject({ row: 18, col: 0 });
  });

  it('records the formula with the subtask, and the NEXT wave does not re-emit it (the step-8 strand)', async () => {
    const run = runWith(monthStates());
    agentRunState.nextExecutableWave.mockReturnValue({ waveIndex: 1, subtasks: [run.subtasks[2]] });
    orchestrator.runStepwiseWave.mockResolvedValue(waveOf('s_main', mainHeaderWave()));

    await exec(run);

    const recorded = agentRunState.recordWaveResult.mock.calls[0][2] as Array<{
      subtaskId: string;
      actions: SheetActionPayload[];
    }>;
    const mainState = recorded.find((s) => s.subtaskId === 's_main')!;
    expect(mainState.actions.some((a) => a.type === 'SET_FORMULA' && a.row === 18 && a.col === 0)).toBe(true);

    // Next wave: formatting only, run carrying what the first wave recorded.
    const next = runWith([...monthStates(), { ...mainState, completed: true }]);
    agentRunState.nextExecutableWave.mockReturnValue({ waveIndex: 2, subtasks: [next.subtasks[3]] });
    orchestrator.runStepwiseWave.mockResolvedValue(
      waveOf('s_fmt', [{ type: 'HIDE_GRIDLINES', sheetName: 'Main' } as SheetActionPayload]),
    );

    await exec(next);

    expect(previewed(1)).toEqual([{ type: 'HIDE_GRIDLINES', sheetName: 'Main' }]);
  });

  it('does not add the formula when an EARLIER wave already wrote inside its spill area', async () => {
    const [addMain, header, ...stray] = mainHeaderWave();
    const run = runWith([
      ...monthStates(),
      { subtaskId: 's_main', completed: true, actions: [addMain, ...stray] },
    ]);
    agentRunState.nextExecutableWave.mockReturnValue({ waveIndex: 2, subtasks: [run.subtasks[3]] });
    orchestrator.runStepwiseWave.mockResolvedValue(waveOf('s_fmt', [header]));

    await exec(run);

    expect(previewed().some((a) => a.type === 'SET_FORMULA')).toBe(false);
  });
});
