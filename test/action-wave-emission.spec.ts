import { ConversationService } from '../src/excel-ai/services/conversation.service';
import { ChangeSetService } from '../src/audit/change-set.service';
import { WorkbookContext } from '../src/agents/types/agent.types';
import { SheetAction } from '../src/excel-ai/types/sheet-actions.types';

/**
 * Unit coverage for ConversationService.createActionWaveChangeSets — the
 * orchestration that turns a flattened, verified action list into one
 * ChangeSet per accept wave. splitIntoActionWaves (the actual split decision)
 * is covered in action-wave.util.spec.ts; this covers the wiring around it:
 * one createPreview call per wave, in wave order, with results paired back
 * to their wave for the caller to chain dependsOnChangeSetId.
 */
describe('ConversationService.createActionWaveChangeSets', () => {
  let service: ConversationService;
  let createPreview: jest.Mock;

  const context: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [],
    namedRanges: [],
    tables: [],
    onDemandFetchEnabled: false,
  };

  beforeEach(() => {
    createPreview = jest.fn(async ({ actions }: { actions: SheetAction[] }) => ({
      changeSetId: `cs_${actions[0]?.type ?? 'empty'}_${Math.random().toString(36).slice(2, 6)}`,
      conversationId: 'conv-1',
      traceId: 'trace-1',
      timestamp: new Date(),
      prompt: 'test',
      beforeState: {},
      changes: actions.map((a) => ({
        cell: 'A1',
        sheet: a.sheetName ?? 'Sheet1',
        before: null,
        after: 'x',
        isHardcoded: false,
      })),
      actions,
      status: 'previewed',
    }));

    const changeSetService = { createPreview } as unknown as ChangeSetService;

    service = new ConversationService(
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      changeSetService,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
    );
  });

  interface WaveResult {
    wave: { actions: SheetAction[]; label: string };
    changeSet: { changeSetId: string; changes: unknown[] };
  }

  type CreateActionWaveChangeSets = (
    actions: SheetAction[],
    input: { conversationId: string; traceId: string; prompt: string; context: WorkbookContext },
  ) => Promise<WaveResult[]>;

  function callCreateWaveChangeSets(actions: SheetAction[]): Promise<WaveResult[]> {
    const method = (service as unknown as Record<string, unknown>)
      .createActionWaveChangeSets as CreateActionWaveChangeSets;
    return method.call(service, actions, {
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'build the ledger',
      context,
    });
  }

  it('creates exactly one ChangeSet for a pure-write batch (no split)', async () => {
    const actions: SheetAction[] = [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 0, col: 0, value: 'x' },
    ];

    const results = await callCreateWaveChangeSets(actions);

    expect(createPreview).toHaveBeenCalledTimes(1);
    expect(results).toHaveLength(1);
    expect(results[0].wave.actions).toEqual(actions);
  });

  // TASKS.md #141 — a mixed create+write batch is ONE ChangeSet and one Accept
  // card now. It used to be two, the second gated behind the first; accepting
  // only the first left the user with created-but-empty sheets.
  it('creates a single ChangeSet for a mixed create+write batch', async () => {
    const actions: SheetAction[] = [
      { type: 'ADD_SHEET', name: 'January' } as SheetAction,
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0, value: 'Unit No' },
    ];

    const results = await callCreateWaveChangeSets(actions);

    expect(createPreview).toHaveBeenCalledTimes(1);
    expect(results).toHaveLength(1);
    expect(createPreview.mock.calls[0][0].actions).toEqual(actions);
  });

  it('passes the whole batch as provenance sourceRefs input, creates first', async () => {
    const actions: SheetAction[] = [
      { type: 'SET_FORMULA', sheetName: 'February', row: 0, col: 1, formula: '=A1' },
      { type: 'ADD_SHEET', name: 'February' } as SheetAction,
    ];

    await callCreateWaveChangeSets(actions);

    const callArgs = createPreview.mock.calls[0][0];
    expect(callArgs.actions).toHaveLength(2);
    // The create is hoisted ahead of the write that depends on it.
    expect(callArgs.actions[0].type).toBe('ADD_SHEET');
    expect(callArgs.actions[1].type).toBe('SET_FORMULA');
  });
});
