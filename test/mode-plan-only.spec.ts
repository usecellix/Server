import { ConversationService } from '../src/excel-ai/services/conversation.service';
import { Tier0DirectService } from '../src/excel-ai/services/tier0-direct.service';
import { Tier2GenerateVerifyService } from '../src/excel-ai/services/tier2-generate-verify.service';
import { OrchestratorService } from '../src/agents/orchestrator.service';
import { ChangeSetService } from '../src/audit/change-set.service';
import { StructuredLogger } from '../src/agents/logging/structured-logger';
import { FormulaAnalyzer } from '../src/formula/formula.analyzer';
import { SheetAnalyzerService } from '../src/excel-ai/services/sheet-analyzer.service';
import { RouterDecision } from '../src/excel-ai/types/router.types';
import { ConversationRequestDto } from '../src/excel-ai/dto/conversation-request.dto';
import { FastifyReply } from 'fastify';
import * as sseUtil from '../src/excel-ai/utils/sse.util';

describe('ConversationService plan mode (streamPlanOnly)', () => {
  let service: ConversationService;
  let orchestrator: { planOnly: jest.Mock; run: jest.Mock };
  let tier2GenerateVerify: { generateOnly: jest.Mock; execute: jest.Mock };
  let changeSetService: { createPreview: jest.Mock };
  let structuredLogger: { logTierDecision: jest.Mock };
  const emittedEvents: Array<{ event: string; data: Record<string, unknown> }> = [];

  const baseRequest: ConversationRequestDto = {
    message: 'sort by column B and create a summary chart',
    sheetData: [
      ['Name', 'Amount'],
      ['Alpha', 100],
      ['Beta', 200],
    ],
    mode: 'plan',
    workbookContext: {
      activeSheet: 'Sheet1',
      sheets: [
        {
          sheetName: 'Sheet1',
          usedRange: 'A1:B3',
          rowCount: 3,
          colCount: 2,
          headers: ['Name', 'Amount'],
        },
      ],
    },
  };

  const reply = {} as FastifyReply;
  const analysis = { rowCount: 3, columnCount: 2, headers: ['Name', 'Amount'] };

  beforeEach(() => {
    emittedEvents.length = 0;
    orchestrator = {
      planOnly: jest.fn(),
      run: jest.fn(),
    };
    tier2GenerateVerify = {
      generateOnly: jest.fn(),
      execute: jest.fn(),
    };
    changeSetService = {
      createPreview: jest.fn(),
    };
    structuredLogger = {
      logTierDecision: jest.fn(),
    };

    const formulaAnalyzer = {
      analyzeSheet: jest.fn().mockReturnValue({}),
    };

    service = new ConversationService(
      {} as never,
      {} as SheetAnalyzerService,
      {} as never,
      {} as never,
      changeSetService as unknown as ChangeSetService,
      {} as never,
      orchestrator as unknown as OrchestratorService,
      {} as never,
      {} as never,
      {} as never,
      {} as never,
      formulaAnalyzer as unknown as FormulaAnalyzer,
      {} as never,
      {} as never,
      new Tier0DirectService(),
      {} as never,
      tier2GenerateVerify as unknown as Tier2GenerateVerifyService,
      structuredLogger as unknown as StructuredLogger,
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

  it('tier 3 plan mode calls planOnly once and never creates a ChangeSet', async () => {
    orchestrator.planOnly.mockResolvedValue({
      subtasks: [
        {
          id: 's1',
          description: 'Sort column B',
          targetSheet: 'Sheet1',
          dependsOn: [],
          estimatedActions: 1,
        },
        {
          id: 's2',
          description: 'Create summary chart',
          targetSheet: 'Sheet1',
          dependsOn: ['s1'],
          estimatedActions: 1,
        },
      ],
      clarificationsNeeded: [],
      confidence: 'high',
      reasoning: 'Two-step plan',
    });

    const routerDecision: RouterDecision = {
      route: 'write',
      confidence: 0.95,
      reasoning: 'compound',
      complexity: 3,
      actionHint: 'COMPOUND',
      matchedBy: 'regex',
    };

    await (service as any).streamPlanOnly(
      baseRequest,
      routerDecision,
      'conv-1',
      'trace-1',
      reply,
      [],
      analysis,
      emit,
    );

    expect(orchestrator.planOnly).toHaveBeenCalledTimes(1);
    expect(orchestrator.run).not.toHaveBeenCalled();
    expect(changeSetService.createPreview).not.toHaveBeenCalled();
    expect(emittedEvents.some((e) => e.event === 'plan_only')).toBe(true);
    expect(emittedEvents.some((e) => e.event === 'actions')).toBe(false);
  });

  it('tier 2 plan mode uses generateOnly and emits plan_only without ChangeSet', async () => {
    tier2GenerateVerify.generateOnly.mockResolvedValue({
      actions: [
        { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 1, formula: '=B2*0.18' },
      ],
      answer: 'Proposed 1 change',
      durationMs: 12,
    });

    const routerDecision: RouterDecision = {
      route: 'write',
      confidence: 0.95,
      reasoning: 'formula',
      complexity: 2,
      actionHint: 'FORMULA_GEN',
      matchedBy: 'regex',
    };

    await (service as any).streamPlanOnly(
      { ...baseRequest, message: 'calculate GST at 18% in column C' },
      routerDecision,
      'conv-2',
      'trace-2',
      reply,
      [],
      analysis,
      emit,
    );

    expect(tier2GenerateVerify.generateOnly).toHaveBeenCalledTimes(1);
    expect(orchestrator.planOnly).not.toHaveBeenCalled();
    expect(changeSetService.createPreview).not.toHaveBeenCalled();
    const planOnly = emittedEvents.find((e) => e.event === 'plan_only');
    expect(planOnly?.data.proposedActions).toHaveLength(1);
  });

  it('tier 0 plan mode describes the action without LLM or ChangeSet', async () => {
    const routerDecision: RouterDecision = {
      route: 'write',
      confidence: 0.95,
      reasoning: 'format',
      complexity: 0,
      actionHint: 'FREEZE_PANES',
      matchedBy: 'regex',
    };

    await (service as any).streamPlanOnly(
      { ...baseRequest, message: 'freeze top row' },
      routerDecision,
      'conv-3',
      'trace-3',
      reply,
      [],
      analysis,
      emit,
    );

    expect(orchestrator.planOnly).not.toHaveBeenCalled();
    expect(tier2GenerateVerify.generateOnly).not.toHaveBeenCalled();
    expect(changeSetService.createPreview).not.toHaveBeenCalled();
    const planOnly = emittedEvents.find((e) => e.event === 'plan_only');
    expect(planOnly?.data.steps).toEqual(
      expect.arrayContaining([expect.objectContaining({ title: expect.stringMatching(/freeze/i) })]),
    );
  });
});
