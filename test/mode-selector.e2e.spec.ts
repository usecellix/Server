/**
 * Spec 08 — mode selector e2e coverage (ask / plan / act).
 * Complements mode-plan-only.spec.ts with branching contracts.
 */
import { normalizeAssistantMode, modeIsReadOnly } from '../src/excel-ai/utils/mode-guard.util';
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

describe('mode-selector.e2e', () => {
  describe('DTO / mode normalization', () => {
    it('defaults omitted mode to act/action', () => {
      expect(normalizeAssistantMode(undefined)).toBe('action');
      expect(normalizeAssistantMode('act')).toBe('action');
    });

    it('marks ask and plan as read-only', () => {
      expect(modeIsReadOnly('ask')).toBe(true);
      expect(modeIsReadOnly('plan')).toBe(true);
      expect(modeIsReadOnly('action')).toBe(false);
      expect(modeIsReadOnly('act')).toBe(false);
    });
  });

  describe('plan mode never creates ChangeSet', () => {
    let service: ConversationService;
    let orchestrator: { planOnly: jest.Mock; run: jest.Mock };
    let tier2GenerateVerify: { generateOnly: jest.Mock };
    let changeSetService: { createPreview: jest.Mock };
    const emitted: Array<{ event: string; data: Record<string, unknown> }> = [];

    const analysis = { rowCount: 3, columnCount: 2, headers: ['Name', 'Amount'] };
    const reply = {} as FastifyReply;

    beforeEach(() => {
      emitted.length = 0;
      orchestrator = { planOnly: jest.fn(), run: jest.fn() };
      tier2GenerateVerify = { generateOnly: jest.fn() };
      changeSetService = { createPreview: jest.fn() };

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
        { analyzeSheet: jest.fn().mockReturnValue({}) } as unknown as FormulaAnalyzer,
        {} as never,
        {} as never,
        new Tier0DirectService(),
        {} as never,
        tier2GenerateVerify as unknown as Tier2GenerateVerifyService,
        { logTierDecision: jest.fn() } as unknown as StructuredLogger,
      );

      jest.spyOn(service as never, 'saveMessage' as never).mockResolvedValue(undefined as never);
      jest.spyOn(service as never, 'markCompleted' as never).mockResolvedValue(undefined as never);
      jest.spyOn(sseUtil, 'endSseResponse').mockImplementation(() => undefined);
    });

    afterEach(() => {
      jest.restoreAllMocks();
    });

    const emit = (event: string, data: Record<string, unknown>) => {
      emitted.push({ event, data });
    };

    const baseRequest: ConversationRequestDto = {
      message: 'sort by column B and create a chart',
      sheetData: [
        ['Name', 'Amount'],
        ['A', 1],
      ],
      mode: 'plan',
      workbookContext: {
        activeSheet: 'Sheet1',
        sheets: [
          {
            sheetName: 'Sheet1',
            usedRange: 'A1:B2',
            rowCount: 2,
            colCount: 2,
            headers: ['Name', 'Amount'],
          },
        ],
      },
    };

    it('ask mode is read-only (no ChangeSet path)', () => {
      expect(modeIsReadOnly(normalizeAssistantMode('ask'))).toBe(true);
      expect(changeSetService.createPreview).not.toHaveBeenCalled();
    });

    it('act mode is the only write-enabled alias', () => {
      expect(normalizeAssistantMode('act')).toBe('action');
      expect(modeIsReadOnly('act')).toBe(false);
    });

    it('plan mode tier 3 emits plan_only and never createPreview', async () => {
      orchestrator.planOnly.mockResolvedValue({
        subtasks: [
          {
            id: 's1',
            description: 'Sort',
            targetSheet: 'Sheet1',
            dependsOn: [],
            estimatedActions: 1,
          },
        ],
        clarificationsNeeded: [],
        confidence: 'high',
        reasoning: 'ok',
      });

      const decision: RouterDecision = {
        route: 'write',
        confidence: 0.9,
        reasoning: 'compound',
        complexity: 3,
        matchedBy: 'regex',
      };

      await (service as any).streamPlanOnly(
        baseRequest,
        decision,
        'conv-mode',
        'trace-mode',
        reply,
        [],
        analysis,
        emit,
      );

      expect(orchestrator.planOnly).toHaveBeenCalledTimes(1);
      expect(orchestrator.run).not.toHaveBeenCalled();
      expect(changeSetService.createPreview).not.toHaveBeenCalled();
      expect(emitted.some((e) => e.event === 'plan_only')).toBe(true);
      expect(emitted.some((e) => e.event === 'actions')).toBe(false);
    });
  });
});
