import { ConversationService } from '../src/excel-ai/services/conversation.service';
import { Tier0DirectService } from '../src/excel-ai/services/tier0-direct.service';
import { RouterDecision } from '../src/excel-ai/types/router.types';
import { ConversationRequestDto } from '../src/excel-ai/dto/conversation-request.dto';
import { FastifyReply } from 'fastify';
import * as sseUtil from '../src/excel-ai/utils/sse.util';

/**
 * Wiring test for the credit gate/debit hooks added to:
 *  - handleWriteRoute's Tier 1 single-action dispatch (a real gap: Tier 1
 *    writes — e.g. "highlight that row in green" — ran a real LLM call and
 *    changed the workbook, with NO gate or debit at all before this)
 *  - handleDataQueryRoute, the new wrapper around handleSmartDataQuery (same
 *    gap for the add-in's "find X" / data-lookup route)
 *
 * Both price as CREDIT_SYSTEM.md CD-1's own table specifies: Tier 1 writes
 * as FORMULA_GENERATE_OR_FIX (same category Tier 2 uses — see
 * credit-write-route-integration.spec.ts), data-query as FORMULA_QA_SIMPLE
 * (the doc's own "Formula Q&A" mapping, same action type web-chat's
 * WebChatService.ask uses for a single-scope question).
 *
 * Same direct-private-method-call harness pattern as
 * credit-write-route-integration.spec.ts.
 */
describe('ConversationService — credit gate/debit around Tier 1 write route', () => {
  let service: ConversationService;
  let tier1SingleAction: { execute: jest.Mock };
  let changeSetService: { createPreview: jest.Mock };
  let creditGate: { checkBalance: jest.Mock; hasAnyBalance: jest.Mock };
  let creditLedger: { debit: jest.Mock };
  const emittedEvents: Array<{ event: string; data: Record<string, unknown> }> = [];

  const baseRequest: ConversationRequestDto = {
    message: 'highlight that row in green',
    sheetData: [
      ['Name', 'Amount'],
      ['Alpha', 100],
    ],
  };

  // Tier 1 dispatch requires complexity === 1 with an actionHint present.
  const routerDecision: RouterDecision = {
    route: 'write',
    confidence: 0.9,
    reasoning: 'conditional format',
    complexity: 1,
    actionHint: 'CONDITIONAL_FORMAT',
    matchedBy: 'regex',
  };

  const reply = {} as FastifyReply;
  const analysis = {
    rowCount: 2,
    columnCount: 2,
    headers: ['Name', 'Amount'],
    isEmpty: false,
    columnLetters: ['A', 'B'],
  };

  const successfulTier1Result = {
    actions: [
      {
        type: 'FORMAT_MATCHING_ROWS',
        sheetName: 'Sheet1',
        range: 'A1:B2',
        hasHeaders: true,
        filter: { column: 'Amount', operator: 'equals', value: 100 },
        format: { fillColor: '#C6EFCE' },
      },
    ],
    answer: "I'll highlight the matching rows in green.",
    model: 'z-ai/glm-5.3',
  };

  beforeEach(() => {
    emittedEvents.length = 0;
    tier1SingleAction = { execute: jest.fn().mockResolvedValue(successfulTier1Result) };
    changeSetService = {
      createPreview: jest.fn().mockResolvedValue({
        changeSetId: 'cs_t1',
        changes: [],
        irreversibleActionTypes: ['FORMAT_MATCHING_ROWS'],
      }),
    };
    creditGate = { checkBalance: jest.fn(), hasAnyBalance: jest.fn().mockResolvedValue(true) };
    creditLedger = { debit: jest.fn() };

    const workflowTrace = {
      startTrace: jest.fn(),
      appendNode: jest.fn(),
      setMeta: jest.fn(),
      finalize: jest.fn(),
    };
    const structuredLogger = { logTierDecision: jest.fn() };
    const formulaAnalyzer = { analyzeSheet: jest.fn().mockReturnValue({}) };
    const engine = { finalizeActions: jest.fn((actions: unknown[]) => actions) };

    service = new ConversationService(
      {} as never,
      {} as never,
      engine as never,
      {} as never,
      changeSetService as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      formulaAnalyzer as never,
      {} as never,
      {} as never,
      new Tier0DirectService(),
      tier1SingleAction as never,
      {} as never,
      structuredLogger as never,
      workflowTrace as never,
      creditGate as never,
      creditLedger as never,
      {} as never,
    );

    jest.spyOn(service as never, 'saveMessage' as never).mockResolvedValue(undefined as never);
    jest.spyOn(service as never, 'markCompleted' as never).mockResolvedValue(undefined as never);
    jest.spyOn(sseUtil, 'endSseResponse').mockImplementation(() => undefined);
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  const emit = (event: string, data: Record<string, unknown>) => {
    emittedEvents.push({ event, data });
  };

  function callHandleWriteRoute(userId?: string) {
    return (service as any).handleWriteRoute(
      baseRequest,
      routerDecision,
      'conv-1',
      'trace-1',
      reply,
      [],
      analysis,
      emit,
      userId,
    );
  }

  it('gates before the Tier 1 LLM call and blocks when balance is insufficient', async () => {
    creditGate.checkBalance.mockResolvedValue({
      allowed: false,
      reason: 'insufficient_balance',
      availableBalance: 3,
      requiredCredits: 8,
    });

    await callHandleWriteRoute('user-1');

    expect(creditGate.checkBalance).toHaveBeenCalledWith('user-1', 'FORMULA_GENERATE_OR_FIX');
    expect(tier1SingleAction.execute).not.toHaveBeenCalled();
    expect(changeSetService.createPreview).not.toHaveBeenCalled();
    expect(creditLedger.debit).not.toHaveBeenCalled();

    const errorEvent = emittedEvents.find((e) => e.event === 'error');
    expect(errorEvent?.data).toEqual(
      expect.objectContaining({ code: 'INSUFFICIENT_CREDIT', availableBalance: 3, requiredCredits: 8 }),
    );
  });

  it('dispatches to Tier 1 and debits once actions stream, when balance allows', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: true, availableBalance: 100, requiredCredits: 8 });
    creditLedger.debit.mockResolvedValue({
      debited: true,
      balances: { planCredits: 92, purchasedCredits: 0, oneTimeCredits: 0 },
    });

    await callHandleWriteRoute('user-1');

    expect(tier1SingleAction.execute).toHaveBeenCalledTimes(1);
    expect(changeSetService.createPreview).toHaveBeenCalledTimes(1);
    expect(creditLedger.debit).toHaveBeenCalledWith('user-1', 'FORMULA_GENERATE_OR_FIX', 1, {
      conversationId: 'conv-1',
    });

    const creditsEvent = emittedEvents.find((e) => e.event === 'credits');
    expect(creditsEvent?.data).toEqual({
      planCredits: 92,
      purchasedCredits: 0,
      oneTimeCredits: 0,
      debited: 8,
      actionType: 'FORMULA_GENERATE_OR_FIX',
    });
  });

  it('never debits when Tier 1 produces no actions (falls through, past this branch entirely)', async () => {
    // An empty-actions Tier 1 result falls all the way through to Tier 3's own
    // gate (this router decision's complexity is 1, not 2, so the Tier 2
    // `else if` branch is skipped too) — that gate is pre-existing,
    // independently tested behavior, not part of what this suite covers.
    // hasAnyBalance -> false ends the turn right there with its own 'error'
    // emit, which is enough to prove the Tier 1 branch itself never debited.
    creditGate.checkBalance.mockResolvedValue({ allowed: true, availableBalance: 100, requiredCredits: 8 });
    creditGate.hasAnyBalance.mockResolvedValue(false);
    tier1SingleAction.execute.mockResolvedValue({
      actions: [],
      answer: 'Needs a bigger plan.',
    });

    await callHandleWriteRoute('user-1');

    expect(changeSetService.createPreview).not.toHaveBeenCalled();
    expect(creditLedger.debit).not.toHaveBeenCalled();
    expect(emittedEvents.find((e) => e.event === 'credits')).toBeUndefined();
  });

  it('skips the gate and debit entirely when userId is absent (eval-bypass path)', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: false, reason: 'insufficient_balance' });

    await callHandleWriteRoute(undefined);

    expect(creditGate.checkBalance).not.toHaveBeenCalled();
    expect(tier1SingleAction.execute).toHaveBeenCalledTimes(1);
    expect(changeSetService.createPreview).toHaveBeenCalledTimes(1);
    expect(creditLedger.debit).not.toHaveBeenCalled();
  });

  it('does not fail an already-streamed turn when the debit loses a balance race (CD-4/CD-6)', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: true, availableBalance: 8, requiredCredits: 8 });
    creditLedger.debit.mockResolvedValue({ debited: false });

    await callHandleWriteRoute('user-1');

    expect(changeSetService.createPreview).toHaveBeenCalledTimes(1);
    expect(emittedEvents.find((e) => e.event === 'credits')).toBeUndefined();
    expect(emittedEvents.find((e) => e.event === 'error')).toBeUndefined();
  });
});

describe('ConversationService — credit gate/debit around handleDataQueryRoute', () => {
  let service: ConversationService;
  let smartDataQuery: { handleQuery: jest.Mock };
  let creditGate: { checkBalance: jest.Mock };
  let creditLedger: { debit: jest.Mock };
  const emittedEvents: Array<{ event: string; data: Record<string, unknown> }> = [];

  const baseRequest: ConversationRequestDto = {
    message: 'find 310',
    sheetData: [
      ['Name', 'Amount'],
      ['Alpha', 310],
    ],
  };

  const analysis = {
    rowCount: 2,
    columnCount: 2,
    headers: ['Name', 'Amount'],
    isEmpty: false,
    columnLetters: ['A', 'B'],
  };

  beforeEach(() => {
    emittedEvents.length = 0;
    smartDataQuery = { handleQuery: jest.fn().mockResolvedValue('Found 310 in row 2.') };
    creditGate = { checkBalance: jest.fn() };
    creditLedger = { debit: jest.fn() };

    const workflowTrace = { startTrace: jest.fn(), appendNode: jest.fn(), setMeta: jest.fn(), finalize: jest.fn() };

    service = new ConversationService(
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
      smartDataQuery as never,
      new Tier0DirectService(),
      {} as never,
      {} as never,
      {} as never,
      workflowTrace as never,
      creditGate as never,
      creditLedger as never,
      {} as never,
    );

    jest.spyOn(service as never, 'saveMessage' as never).mockResolvedValue(undefined as never);
    jest.spyOn(service as never, 'markCompleted' as never).mockResolvedValue(undefined as never);
    // emitLocalDecision's 'answer' branch calls buildAnswerPersistMetadata ->
    // getRecentMessages -> conversationModel.findOne — not relevant to what
    // this suite is testing (credit gate/debit), so short-circuit it rather
    // than wire a fake Mongoose model.
    jest.spyOn(service as never, 'buildAnswerPersistMetadata' as never).mockResolvedValue(undefined as never);
    // resolveFindPointers touches sheetAnalyzer/dataQuery for "find X"-shaped
    // messages (this suite's baseRequest is "find 310") — pre-existing,
    // unrelated logic; short-circuit it the same way for the same reason.
    jest.spyOn(service as never, 'resolveFindPointers' as never).mockReturnValue({} as never);
    jest.spyOn(sseUtil, 'endSseResponse').mockImplementation(() => undefined);
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  const emit = (event: string, data: Record<string, unknown>) => {
    emittedEvents.push({ event, data });
  };

  function callHandleDataQueryRoute(userId?: string) {
    return (service as any).handleDataQueryRoute(
      baseRequest,
      analysis,
      'conv-1',
      'trace-1',
      emit,
      userId,
    );
  }

  it('gates before the data-query LLM call and blocks when balance is insufficient', async () => {
    creditGate.checkBalance.mockResolvedValue({
      allowed: false,
      reason: 'insufficient_balance',
      availableBalance: 1,
      requiredCredits: 2,
    });

    await callHandleDataQueryRoute('user-1');

    expect(creditGate.checkBalance).toHaveBeenCalledWith('user-1', 'FORMULA_QA_SIMPLE');
    expect(smartDataQuery.handleQuery).not.toHaveBeenCalled();
    expect(creditLedger.debit).not.toHaveBeenCalled();

    const errorEvent = emittedEvents.find((e) => e.event === 'error');
    expect(errorEvent?.data).toEqual(
      expect.objectContaining({ code: 'INSUFFICIENT_CREDIT', availableBalance: 1, requiredCredits: 2 }),
    );
  });

  it('runs the query and debits FORMULA_QA_SIMPLE once an answer exists, when balance allows', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: true, availableBalance: 100, requiredCredits: 2 });
    creditLedger.debit.mockResolvedValue({
      debited: true,
      balances: { planCredits: 98, purchasedCredits: 0, oneTimeCredits: 0 },
    });

    await callHandleDataQueryRoute('user-1');

    expect(smartDataQuery.handleQuery).toHaveBeenCalledTimes(1);
    expect(creditLedger.debit).toHaveBeenCalledWith('user-1', 'FORMULA_QA_SIMPLE', 1, {
      conversationId: 'conv-1',
    });

    const creditsEvent = emittedEvents.find((e) => e.event === 'credits');
    expect(creditsEvent?.data).toEqual({
      planCredits: 98,
      purchasedCredits: 0,
      oneTimeCredits: 0,
      debited: 2,
      actionType: 'FORMULA_QA_SIMPLE',
    });
  });

  it('skips the gate and debit entirely when userId is absent (eval-bypass path)', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: false, reason: 'insufficient_balance' });

    await callHandleDataQueryRoute(undefined);

    expect(creditGate.checkBalance).not.toHaveBeenCalled();
    expect(smartDataQuery.handleQuery).toHaveBeenCalledTimes(1);
    expect(creditLedger.debit).not.toHaveBeenCalled();
  });

  it('does not fail an already-answered turn when the debit loses a balance race (CD-4/CD-6)', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: true, availableBalance: 2, requiredCredits: 2 });
    creditLedger.debit.mockResolvedValue({ debited: false });

    await callHandleDataQueryRoute('user-1');

    expect(smartDataQuery.handleQuery).toHaveBeenCalledTimes(1);
    expect(emittedEvents.find((e) => e.event === 'credits')).toBeUndefined();
    expect(emittedEvents.find((e) => e.event === 'error')).toBeUndefined();
  });
});
