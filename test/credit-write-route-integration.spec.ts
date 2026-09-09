import { ConversationService } from '../src/excel-ai/services/conversation.service';
import { Tier0DirectService } from '../src/excel-ai/services/tier0-direct.service';
import { RouterDecision } from '../src/excel-ai/types/router.types';
import { ConversationRequestDto } from '../src/excel-ai/dto/conversation-request.dto';
import { FastifyReply } from 'fastify';
import * as sseUtil from '../src/excel-ai/utils/sse.util';

/**
 * Wiring test for the credit gate/debit hook added to handleWriteRoute's
 * Tier 2 dispatch (CREDIT_SYSTEM.md CD-3/CD-4). Everything else about Tier 2
 * is already covered by tier2-generate-verify.service.spec.ts — this suite
 * only exercises the new gate-before-dispatch and debit-after-ChangeSet
 * behavior, using the same direct-private-method-call pattern as
 * mode-plan-only.spec.ts.
 */
describe('ConversationService — credit gate/debit around Tier 2 write route', () => {
  let service: ConversationService;
  let tier2GenerateVerify: { execute: jest.Mock };
  let changeSetService: { createPreview: jest.Mock };
  let creditGate: { checkBalance: jest.Mock };
  let creditLedger: { debit: jest.Mock };
  const emittedEvents: Array<{ event: string; data: Record<string, unknown> }> = [];

  const baseRequest: ConversationRequestDto = {
    message: 'fix the SUM formula in B2',
    sheetData: [
      ['Name', 'Amount'],
      ['Alpha', 100],
    ],
  };

  const routerDecision: RouterDecision = {
    route: 'write',
    confidence: 0.9,
    reasoning: 'formula fix',
    complexity: 2,
    actionHint: 'FORMULA_GEN',
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

  const successfulTier2Result = {
    actions: [{ type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 1, formula: '=SUM(A1:A1)' }],
    answer: 'Fixed the formula.',
    verifierPassed: true,
    durationMs: 500,
    sourceRefs: [],
    retried: false,
    toolFollowUp: false,
  };

  beforeEach(() => {
    emittedEvents.length = 0;
    tier2GenerateVerify = { execute: jest.fn().mockResolvedValue(successfulTier2Result) };
    changeSetService = {
      createPreview: jest.fn().mockResolvedValue({
        changeSetId: 'cs_1',
        changes: [{ cell: 'B1', sheet: 'Sheet1', before: null, after: '=SUM(A1:A1)', isHardcoded: false }],
        irreversibleActionTypes: [],
      }),
    };
    creditGate = { checkBalance: jest.fn() };
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
      {} as never,
      tier2GenerateVerify as never,
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

  it('gates before the Tier 2 LLM call and blocks when balance is insufficient', async () => {
    creditGate.checkBalance.mockResolvedValue({
      allowed: false,
      reason: 'insufficient_balance',
      availableBalance: 3,
      requiredCredits: 8,
    });

    await callHandleWriteRoute('user-1');

    expect(creditGate.checkBalance).toHaveBeenCalledWith('user-1', 'FORMULA_GENERATE_OR_FIX');
    expect(tier2GenerateVerify.execute).not.toHaveBeenCalled();
    expect(changeSetService.createPreview).not.toHaveBeenCalled();
    expect(creditLedger.debit).not.toHaveBeenCalled();

    const errorEvent = emittedEvents.find((e) => e.event === 'error');
    expect(errorEvent?.data).toEqual(
      expect.objectContaining({ code: 'INSUFFICIENT_CREDIT', availableBalance: 3, requiredCredits: 8 }),
    );
  });

  it('dispatches to Tier 2 and debits once a ChangeSet is created, when balance allows', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: true, availableBalance: 100, requiredCredits: 8 });
    creditLedger.debit.mockResolvedValue({
      debited: true,
      balances: { planCredits: 92, purchasedCredits: 0, oneTimeCredits: 0 },
    });

    await callHandleWriteRoute('user-1');

    expect(tier2GenerateVerify.execute).toHaveBeenCalledTimes(1);
    expect(changeSetService.createPreview).toHaveBeenCalledTimes(1);
    expect(creditLedger.debit).toHaveBeenCalledWith('user-1', 'FORMULA_GENERATE_OR_FIX', 1, {
      conversationId: 'conv-1',
      changeSetId: 'cs_1',
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

  it('never debits when Tier 2 verification fails, even though the gate allowed dispatch', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: true, availableBalance: 100, requiredCredits: 8 });
    tier2GenerateVerify.execute.mockResolvedValue({ ...successfulTier2Result, verifierPassed: false });

    await callHandleWriteRoute('user-1');

    expect(changeSetService.createPreview).not.toHaveBeenCalled();
    expect(creditLedger.debit).not.toHaveBeenCalled();
    expect(emittedEvents.find((e) => e.event === 'credits')).toBeUndefined();
  });

  it('skips the gate and debit entirely when userId is absent (eval-bypass path)', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: false, reason: 'insufficient_balance' });

    await callHandleWriteRoute(undefined);

    expect(creditGate.checkBalance).not.toHaveBeenCalled();
    expect(tier2GenerateVerify.execute).toHaveBeenCalledTimes(1);
    expect(changeSetService.createPreview).toHaveBeenCalledTimes(1);
    expect(creditLedger.debit).not.toHaveBeenCalled();
  });

  it('does not fail an already-verified turn when the debit loses a balance race (CD-4/CD-6)', async () => {
    creditGate.checkBalance.mockResolvedValue({ allowed: true, availableBalance: 8, requiredCredits: 8 });
    creditLedger.debit.mockResolvedValue({ debited: false });

    await callHandleWriteRoute('user-1');

    expect(changeSetService.createPreview).toHaveBeenCalledTimes(1);
    const actionsEvent = emittedEvents.find((e) => e.event === 'actions');
    expect(actionsEvent).toBeDefined();
    expect(emittedEvents.find((e) => e.event === 'credits')).toBeUndefined();
    // No 'error' event either — a lost race is not surfaced to the user as a failure.
    expect(emittedEvents.find((e) => e.event === 'error')).toBeUndefined();
  });
});
