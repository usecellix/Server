import {
  BadRequestException,
  GoneException,
  Injectable,
  Logger,
  NotFoundException,
} from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { FastifyReply } from 'fastify';
import { Model } from 'mongoose';
import { ConversationRequestDto } from '../dto/conversation-request.dto';
import {
  CONVERSATION_TTL_MS,
  Conversation,
  ConversationDocument,
  ConversationMessageEntry,
} from '../schemas/conversation.schema';
import { endSseResponse, initSseResponse, writeSseEvent } from '../utils/sse.util';
import { AuditService } from '../../audit/audit.service';
import {
  getComplexityTieringMode,
  resolveExecutableTier,
} from '../utils/complexity-tiering-flag.util';
import { ChangeSetService } from '../../audit/change-set.service';
import { buildWorkbookSourceRefsFromActions } from '../../audit/utils/provenance.util';
import { ChangeSetRecord } from '../../audit/types/change-set.types';
import { ActionWave, splitIntoActionWaves } from '../utils/action-wave.util';
import { classifyIntent, detectAmbiguity } from '../llm/ambiguity-detector';
import { LLMTier, SheetSnapshot } from '../../types/cellix.types';
import { OrchestratorService } from '../../agents/orchestrator.service';
import { AgentRunStateService, WaveDecision } from '../../agents/agent-run-state.service';
import { AgentRunDocument } from '../../agents/schemas/agent-run.schema';
import { ContinueRunDto } from '../dto/continue-run.dto';
import {
  isStepwiseExecutionEnabled,
  shouldRunStepwise,
} from '../utils/stepwise-execution-flag.util';
import { SseEmitter } from '../../agents/sse.emitter';
import { ToolBridgeService } from '../../agents/tool-bridge.service';
import { buildAgentWorkbookContext } from '../../agents/utils/workbook-context.builder';
import { computeExecutionWaves } from '../../agents/utils/task-graph.util';
import { WriteRouteNoActionError } from '../errors/write-route-no-action.error';
import { CreditGateService } from '../../credit/credit-gate.service';
import { CreditLedgerService } from '../../credit/credit-ledger.service';
import { resolveCreditCost } from '../../credit/credit-cost-catalog';
import { ChitchatService } from './chitchat.service';
import { ConversationEngineService, EngineResponse, LlmRequestError } from './conversation-engine.service';
import { DataQueryService } from './data-query.service';
import { FindExportService, FindExportSheetSlice } from './find-export.service';
import { ContextCacheService } from '../../common/cache/context-cache.service';
import { LlmRouterService } from './llm-router.service';
import { LlmCallTelemetry, OpenRouterService } from './openrouter.service';
import { RouterDecision, RouterInput } from '../types/router.types';
import { buildTieredToon } from '../utils/tiered-toon.util';
import {
  ASK_MODE_READONLY_DIRECTIVE,
  PLAN_MODE_DIRECTIVE,
} from '../prompt/cellix-system-prompt';
import { modeIsReadOnly, normalizeAssistantMode, stripWriteActions } from '../utils/mode-guard.util';
import { PlannerOutput } from '../../agents/types/agent.types';
import { buildStatusMessage } from '../utils/status-message.util';
import { tryDeterministicTableCreate } from '../utils/table-request.util';
import { routeShortcutAction, buildShortcutAnswer } from '../utils/shortcut-router.util';
import {
  buildDeleteSheetAnswer,
  tryLocalDeleteSheetActions,
} from '../utils/local-sheet-actions.util';
import { stripSheetMentions } from '../utils/sheet-mentions.util';
import {
  resolveConversationHistory,
  resolveEngineWorkbookMeta,
  resolveWorkbookContext,
} from '../utils/workbook-context-resolver.util';
import { buildRefinementContext } from '../utils/refinement-context.util';
import {
  collectRecentTurnActionRecords,
  extractTurnActionRecords,
  formatTurnActionRecordsForExecutor,
  referencesPriorChartOrTable,
} from '../utils/turn-action-history.util';
import { annotateExplicitOverwriteConfirmation } from '../utils/overwrite-confirmation.util';
import { buildEnrichedPromptContext } from '../../formula/enrich-context.util';
import { FormulaAnalyzer } from '../../formula/formula.analyzer';
import { SmartDataQueryService } from './smart-data-query.service';
import { SheetAnalyzerService } from './sheet-analyzer.service';
import { Tier0DirectService, Tier0Result } from './tier0-direct.service';
import { Tier1SingleActionService } from './tier1-single-action.service';
import { Tier2GenerateVerifyService } from './tier2-generate-verify.service';
import { assessTierEscalation } from '../utils/tier-escalation.util';
import { attributeActionsToSubtasks, intentForWave } from '../utils/wave-intent.util';
import {
  selectEarlyEmittable,
  splitEarlyByPhase,
  excludeAlreadyEmitted,
  keysFor,
} from '../utils/progressive-emit.util';
import { StructuredLogger } from '../../agents/logging/structured-logger';
import { WorkbookContext as AgentWorkbookContext } from '../../agents/types/agent.types';
import { SheetAction } from '../types/sheet-actions.types';
import { isFindLookupMessage } from '../utils/find-query-parser.util';
import {
  buildInternalDetails,
  buildUserFacingSummary,
  sanitizeAnswerForUser,
  tierProcessingLabel,
} from '../utils/user-facing-response.util';
import {
  buildSheetOverview,
  formatSheetOverviewMarkdown,
  isSheetOverviewRequest,
  sanitizeAskAnswer,
} from '../utils/sheet-overview.util';
import {
  buildPendingWritePlanMetadata,
  buildResumedWritePrompt,
  findPendingWritePlan,
  isAffirmationMessage,
  shouldStorePendingWritePlan,
} from '../utils/pending-write-plan.util';
import {
  classifyLlmFailure,
  describeLlmFailureForStatus,
  type LlmFailure,
} from '../utils/llm-failure-message.util';
import { annotateAnswerConsistency } from '../utils/answer-consistency.util';
import { deriveConversationTitle, truncateTitle } from '../utils/conversation-title.util';
import { WorkflowTraceService } from '../../common/logging/workflow-trace.service';
import type { WorkflowTraceStatus } from '../../common/logging/schemas/workflow-trace.schema';

const MAX_MESSAGES = 50;

/** History list page size ceiling (TASKS.md #171) — never return unbounded history. */
const HISTORY_MAX_LIMIT = 50;
const HISTORY_DEFAULT_LIMIT = 25;

/** One row of the history list — summary only, no message bodies (TASKS.md #171). */
export interface ConversationSummary {
  conversationId: string;
  workbookId?: string;
  title: string;
  firstMessage: string;
  lastMessage: string;
  messageCount: number;
  status: string;
  updatedAt: Date | null;
}

@Injectable()
export class ConversationService {
  private readonly logger = new Logger(ConversationService.name);

  constructor(
    @InjectModel(Conversation.name)
    private readonly conversationModel: Model<ConversationDocument>,
    private readonly sheetAnalyzer: SheetAnalyzerService,
    private readonly engine: ConversationEngineService,
    private readonly auditService: AuditService,
    private readonly changeSetService: ChangeSetService,
    private readonly openRouter: OpenRouterService,
    private readonly orchestrator: OrchestratorService,
    private readonly llmRouter: LlmRouterService,
    private readonly chitchat: ChitchatService,
    private readonly contextCache: ContextCacheService,
    private readonly dataQuery: DataQueryService,
    private readonly findExport: FindExportService,
    private readonly formulaAnalyzer: FormulaAnalyzer,
    private readonly toolBridge: ToolBridgeService,
    private readonly smartDataQuery: SmartDataQueryService,
    private readonly tier0Direct: Tier0DirectService,
    private readonly tier1SingleAction: Tier1SingleActionService,
    private readonly tier2GenerateVerify: Tier2GenerateVerifyService,
    private readonly structuredLogger: StructuredLogger,
    private readonly workflowTrace: WorkflowTraceService,
    private readonly creditGate: CreditGateService,
    private readonly creditLedger: CreditLedgerService,
    private readonly agentRunState: AgentRunStateService,
  ) {}

  private enrichAgentContext(
    context: AgentWorkbookContext,
    basePromptContext?: string,
    history?: ConversationMessageEntry[],
    userMessage?: string,
  ): { enrichedContext: AgentWorkbookContext; promptContext: string } {
    // Perf #70: analyzeSheet() walks every formula cell on a sheet to build its
    // llmSummary — previously ran unconditionally for EVERY sheet in the workbook
    // context on every Tier 2/3 request, even sheets the request never touches
    // (e.g. one of 16 monthly sheets in a hospitality-workbook build where only
    // "Main" needs formula insight). Scope to sheets that are actually relevant:
    // the active sheet, plus any sheet explicitly named in the user's message
    // (covers cross-sheet formula requests like "fix the SUMIF on January").
    const relevantSheetNames = this.resolveFormulaRelevantSheets(context, userMessage);
    const enrichedSheets = context.sheets.map((sheet) =>
      relevantSheetNames.has(sheet.name)
        ? { ...sheet, formulaInsights: this.formulaAnalyzer.analyzeSheet(sheet) }
        : sheet,
    );
    let enrichedContext: AgentWorkbookContext = { ...context, sheets: enrichedSheets };

    if (history?.length) {
      const priorTurnActions = collectRecentTurnActionRecords(history);
      if (priorTurnActions.length > 0) {
        const summary = formatTurnActionRecordsForExecutor(priorTurnActions);
        enrichedContext = {
          ...enrichedContext,
          priorTurnActions,
          priorTurnActionsSummary:
            userMessage && referencesPriorChartOrTable(userMessage)
              ? `${summary}\nThis follow-up references prior chart/table context — reuse the sourceRange/chartId above for "the current" / "same data"; do not invent a different range.`
              : summary,
        };
      }
    }

    const promptContext = buildEnrichedPromptContext(basePromptContext, enrichedSheets);
    return { enrichedContext, promptContext };
  }

  /**
   * Perf #70: scope formula analysis to sheets the request can plausibly touch —
   * the active sheet, always, plus any other sheet named verbatim in the user's
   * message (so "fix the formula in December" still gets December analyzed even
   * though it isn't active). Deliberately conservative: a large multi-sheet
   * workbook (e.g. 12+ month sheets) with a request scoped to one or two sheets
   * no longer pays analyzeSheet()'s full-formula-walk cost for every other sheet.
   */
  private resolveFormulaRelevantSheets(
    context: AgentWorkbookContext,
    userMessage?: string,
  ): Set<string> {
    const relevant = new Set<string>([context.activeSheetName]);
    if (!userMessage) return relevant;

    for (const sheet of context.sheets) {
      if (sheet.name && userMessage.includes(sheet.name)) {
        relevant.add(sheet.name);
      }
    }
    return relevant;
  }

  private buildWriteMetadata(
    actions: SheetAction[],
    changeSetId?: string,
    extra?: ConversationMessageEntry['metadata'],
  ): ConversationMessageEntry['metadata'] {
    const turnActionRecords = extractTurnActionRecords(actions);
    return {
      actions,
      ...(changeSetId ? { changeSetId } : {}),
      ...(turnActionRecords.length > 0 ? { turnActionRecords } : {}),
      ...extra,
    };
  }

  /**
   * When the assistant only offers a large write ("want me to apply?") with no actions,
   * persist a pendingWritePlan so a short "yes" forces the write path next turn.
   */
  private async buildAnswerPersistMetadata(
    conversationId: string,
    answer: string,
    hasActions: boolean,
  ): Promise<ConversationMessageEntry['metadata'] | undefined> {
    if (!shouldStorePendingWritePlan(answer, hasActions)) {
      return undefined;
    }
    const history = await this.getRecentMessages(conversationId);
    const originalPrompt = this.resolveSubstantiveUserPrompt(history);
    if (!originalPrompt) {
      return undefined;
    }
    return buildPendingWritePlanMetadata(originalPrompt, answer);
  }

  private resolveSubstantiveUserPrompt(
    history: ConversationMessageEntry[],
  ): string | undefined {
    for (let i = history.length - 1; i >= 0; i -= 1) {
      const entry = history[i];
      if (entry?.role !== 'user') continue;
      const content = entry.content?.trim() ?? '';
      if (!content || isAffirmationMessage(content) || content.length < 15) continue;
      return content;
    }
    return undefined;
  }

  private startWorkflowTrace(params: {
    traceId: string;
    conversationId: string;
    workbookId?: string;
    message: string;
    mode?: string;
    request: ConversationRequestDto;
  }): void {
    this.workflowTrace.startTrace({
      traceId: params.traceId,
      conversationId: params.conversationId,
      workbookId: params.workbookId,
      message: params.message,
      mode: params.mode,
      requestInput: {
        message: params.request.message,
        mode: params.mode,
        conversationId: params.conversationId,
        sheetData: params.request.sheetData,
        hasWorkbookContext: Boolean(params.request.workbookContext),
        hasPromptContext: Boolean(params.request.promptContext),
      },
    });
  }

  private logWorkflowRouter(
    traceId: string,
    decision: RouterDecision,
  ): void {
    this.workflowTrace.appendNode(traceId, {
      id: 'router',
      type: 'router',
      label: `Router → ${decision.route}`,
      status: 'success',
      input: { message: decision.reasoning },
      output: {
        route: decision.route,
        confidence: decision.confidence,
        reasoning: decision.reasoning,
        complexity: decision.complexity,
        actionHint: decision.actionHint,
        matchedBy: decision.matchedBy,
        assumption: decision.assumption,
      },
      meta: { route: decision.route },
    });
    this.workflowTrace.setMeta(traceId, {
      route: decision.route,
      ...(decision.complexity !== undefined ? { tier: decision.complexity } : {}),
    });
  }

  private logWorkflowTier(
    traceId: string,
    tier: number,
    actionHint?: string,
    extra?: Record<string, unknown>,
  ): void {
    this.workflowTrace.appendNode(traceId, {
      id: `tier_${tier}`,
      type: 'tier',
      label: `Tier ${tier}`,
      status: 'success',
      input: { actionHint, ...extra },
      output: { tier, actionHint },
      meta: { tier, actionHint },
    });
    this.workflowTrace.setMeta(traceId, { tier });
  }

  private logWorkflowChangeSet(
    traceId: string,
    changeSetId: string,
    actions: SheetAction[],
    changesLength?: number,
  ): void {
    this.workflowTrace.appendNode(traceId, {
      id: 'changeset',
      type: 'changeset',
      label: 'ChangeSet Preview',
      status: 'success',
      input: {
        actionCount: actions.length,
        actionTypes: actions.map((a) => a.type),
      },
      output: {
        changeSetId,
        changesLength,
        actions: actions.map((a) => ({ type: a.type, sheetName: (a as { sheetName?: string }).sheetName })),
      },
      meta: { changeSetId },
    });
    this.workflowTrace.setMeta(traceId, { changeSetId });
  }

  /**
   * Create one ChangeSet per accept wave (see splitIntoActionWaves). Pure
   * orchestration — no SSE emission here, so callers control emit ordering
   * (e.g. 'answer' before the first 'actions' event) exactly as before.
   */
  /**
   * Label for a progressive card — TASKS.md #174.
   *
   * Deliberately plainer than `describeStep` in action-wave.util.ts: mid-run we
   * know what this wave did but not where it sits in the finished build, so the
   * label states the work and claims nothing about position.
   */
  private describeProgressiveWave(actions: SheetAction[]): string {
    const creates = actions.filter((a) =>
      ['ADD_SHEET', 'CREATE_SHEET', 'COPY_SHEET'].includes(String(a.type)),
    ).length;
    if (creates > 0 && creates === actions.length) {
      return `Create ${creates} sheet${creates === 1 ? '' : 's'}`;
    }

    const sheets = new Set(
      actions.map((a) => String(a.sheetName ?? '').trim()).filter(Boolean),
    );
    if (creates > 0) {
      return sheets.size > 1
        ? `Create and fill ${sheets.size} sheets`
        : `Create and fill ${[...sheets][0] || 'sheet'}`;
    }
    return sheets.size > 1
      ? `Write content on ${sheets.size} sheets`
      : `Write content on ${[...sheets][0] || 'sheet'}`;
  }

  private async createActionWaveChangeSets(
    actions: SheetAction[],
    input: { conversationId: string; traceId: string; prompt: string; context: AgentWorkbookContext },
  ): Promise<Array<{ wave: ActionWave; changeSet: ChangeSetRecord }>> {
    const waves = splitIntoActionWaves(actions);
    const activeSheetName = input.context.activeSheetName;
    const results: Array<{ wave: ActionWave; changeSet: ChangeSetRecord }> = [];

    for (const wave of waves) {
      const changeSet = await this.changeSetService.createPreview({
        conversationId: input.conversationId,
        traceId: input.traceId,
        prompt: input.prompt,
        context: input.context,
        actions: wave.actions,
        provenance: {
          sourceRefs: buildWorkbookSourceRefsFromActions(
            wave.actions,
            activeSheetName || 'workbook',
            activeSheetName,
          ),
          workbookId: activeSheetName || 'workbook',
          activeSheetName,
        },
      });
      results.push({ wave, changeSet });
    }

    return results;
  }

  private finalizeWorkflow(
    traceId: string,
    status: WorkflowTraceStatus,
    opts?: {
      durationMs?: number;
      changeSetId?: string;
      route?: string;
      tier?: number;
      sseOutput?: unknown;
    },
  ): void {
    this.workflowTrace.finalize(traceId, {
      status,
      durationMs: opts?.durationMs,
      changeSetId: opts?.changeSetId,
      route: opts?.route,
      tier: opts?.tier,
      sseOutput: opts?.sseOutput,
    });
  }

  /**
   * @param userId Owner of this conversation, resolved from the session by the
   *   controller (TASKS.md #170). Deliberately a separate parameter rather than
   *   a DTO field — a client-supplied userId would let a caller write into, and
   *   later list, another user's history.
   */
  async handleConversation(
    request: ConversationRequestDto,
    reply: FastifyReply,
    traceId = '-',
    userId?: string,
  ): Promise<void> {
    this.validateRequest(request);

    const conversation = await this.getOrCreateConversation(
      request.conversationId,
      request.workbookId,
      userId,
    );
    const activeRequestRaw = await this.applyRefinementContext(request);
    let activeRequest: ConversationRequestDto = {
      ...activeRequestRaw,
      mode: normalizeAssistantMode(activeRequestRaw.mode),
    };
    const requestMode = activeRequest.mode ?? 'action';
    const writeAllowed = requestMode === 'action';
    /** True when this turn is a short "yes" resuming a stored multi-sheet / large write plan. */
    let resumePendingWrite = false;

    // Spec 09 item 3: instant shortcut before SheetAnalyzer (no sheet analysis needed).
    const instantShortcut = this.llmRouter.peekInstantShortcut(activeRequest.message);
    if (instantShortcut && writeAllowed) {
      const activeSheetName = this.resolveActiveSheetName(activeRequest);
      const shortcutActions = routeShortcutAction(activeRequest.message, activeSheetName);
      if (shortcutActions?.length) {
        initSseResponse(reply);
        const conversationId = conversation.conversationId;
        const emit = (event: string, data: Record<string, unknown>) =>
          writeSseEvent(reply, event, { ...data, conversationId });

        this.startWorkflowTrace({
          traceId,
          conversationId,
          workbookId: conversation.workbookId,
          message: activeRequest.message,
          mode: requestMode,
          request: activeRequest,
        });
        this.workflowTrace.appendNode(traceId, {
          id: 'router',
          type: 'router',
          label: 'Router → shortcut',
          status: 'success',
          output: {
            route: 'shortcut',
            confidence: 1,
            reasoning: 'Matched instant shortcut regex — no LLM needed',
          },
        });
        this.workflowTrace.setMeta(traceId, { route: 'shortcut', tier: 0 });

        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}`,
          role: 'user',
          content: request.message,
          type: 'command',
          timestamp: new Date(),
        });

        this.logger.log(
          `[${traceId}] Router: route=shortcut confidence=1 (pre-analyze) "Matched instant shortcut regex — no LLM needed"`,
        );
        emit('status', { message: 'Working on your request…' });

        await this.emitLocalDecision(
          conversationId,
          {
            kind: 'actions',
            answer: buildShortcutAnswer(shortcutActions),
            explanation: 'Matched instant shortcut regex — no LLM needed',
            actions: shortcutActions,
          },
          emit,
          { traceId, route: 'shortcut', tier: 0 },
        );
        endSseResponse(reply);
        return;
      }
      // Regex matched but handler returned null — fall through to full path.
    }

    // CHITCHAT gate — before SheetAnalyzer, same reason as the instant shortcut
    // above: a greeting needs no workbook context, no TOON compression, and no
    // Tier 0-3 dispatch. Fails open to TASK (classifyIntent never throws out of
    // this call), so a classifier outage just costs one extra tiering pass.
    const intentLabel = await this.llmRouter.classifyIntent(activeRequest.message);
    if (intentLabel === 'CHITCHAT') {
      await this.handleChitchat(activeRequest, request, conversation, reply, traceId);
      return;
    }

    let analysis = this.sheetAnalyzer.analyze(activeRequest.sheetData);
    const declaredRowCount = this.resolveDeclaredRowCount(activeRequest);
    if (declaredRowCount > analysis.rowCount) {
      analysis = { ...analysis, rowCount: declaredRowCount };
    }
    const startedAt = Date.now();

    await this.saveMessage(conversation.conversationId, {
      id: `msg_${Date.now()}`,
      role: 'user',
      content: request.message,
      type: 'command',
      timestamp: new Date(),
    });

    conversation.sheetSnapshot = {
      rowCount: analysis.rowCount,
      columnCount: analysis.columnCount,
      headers: analysis.headers,
    };
    await this.conversationModel.updateOne(
      { conversationId: conversation.conversationId },
      {
        $set: {
          sheetSnapshot: conversation.sheetSnapshot,
        },
      },
    );

    initSseResponse(reply);

    const conversationId = conversation.conversationId;
    const emit = (event: string, data: Record<string, unknown>) =>
      writeSseEvent(reply, event, { ...data, conversationId });
    let localReason = this.engine.hasOpenAi() ? 'llm_not_used' : 'no_llm_provider';
    /** Set when an LLM call failed, so local copy can name the real cause (F11). */
    let llmFailure: LlmFailure | undefined;

    this.startWorkflowTrace({
      traceId,
      conversationId,
      workbookId: conversation.workbookId,
      message: activeRequest.message,
      mode: requestMode,
      request: activeRequest,
    });

    try {
      this.logger.log(
        `Conversation request trace=${traceId} conversation=${conversationId} message="${this.clipForLog(activeRequest.message)}" sheet=${analysis.rowCount}x${analysis.columnCount} history=${activeRequest.context?.previousMessages?.length ?? 0}${activeRequest.refinementChangeSetId ? ' quickEdit=true' : ''}`,
      );
      emit('status', {
        message: activeRequest.refinementChangeSetId
          ? 'Quick edit — refining your last change…'
          : buildStatusMessage(activeRequest.message, analysis),
      });

      const history = await this.getRecentMessages(conversation.conversationId);

      // Affirmation resume: rehydrate turn-1 scaffold confirm into a forced write prompt.
      const pendingWritePlan = findPendingWritePlan(history);
      if (
        writeAllowed &&
        isAffirmationMessage(request.message) &&
        pendingWritePlan
      ) {
        resumePendingWrite = true;
        this.logger.log(
          `Pending write plan resume trace=${traceId} conversation=${conversationId} promptChars=${pendingWritePlan.originalPrompt.length}`,
        );
        emit('status', { message: 'Applying your confirmed changes…' });
        activeRequest = {
          ...activeRequest,
          message: buildResumedWritePrompt(pendingWritePlan),
        };
      }

      const deterministicTable = writeAllowed
        ? tryDeterministicTableCreate(request.message)
        : null;
      if (deterministicTable) {
        const { plan, actions } = deterministicTable;
        this.logger.log(
          `Table create (deterministic) trace=${traceId} conversation=${conversationId} rows=${plan.rowCount} cols=${plan.headers.length}`,
        );
        const decision = {
          kind: 'actions' as const,
          answer: `Created **${plan.rowCount}** rows with columns: ${plan.headers.join(', ')}.`,
          explanation: 'Wrote headers and all data rows to your sheet.',
          actions,
        };
        await this.emitLocalDecision(conversation.conversationId, decision, emit, {
          traceId,
          route: 'write',
          tier: 0,
        });
        endSseResponse(reply);
        return;
      }

      const recentHistory = history
        .filter((entry) => entry.role === 'user')
        .slice(-2)
        .map((entry) => entry.content);

      let routerDecision = await this.llmRouter.route(
        this.buildRouterInput(activeRequest, recentHistory, analysis),
      );

      if (resumePendingWrite) {
        routerDecision = {
          ...routerDecision,
          route: 'write',
          complexity: Math.max(routerDecision.complexity ?? 0, 3) as 0 | 1 | 2 | 3,
          confidence: Math.max(routerDecision.confidence, 0.9),
          reasoning: `pending_write_plan_resume: ${routerDecision.reasoning}`,
          overridden: true,
        };
      }

      this.logger.log(
        `[${traceId}] Router: route=${routerDecision.route} confidence=${routerDecision.confidence} "${routerDecision.reasoning}"`,
      );
      this.logWorkflowRouter(traceId, routerDecision);

      const routedRequest = this.applyRoutedPromptContext(
        activeRequest,
        conversationId,
        routerDecision,
        analysis,
        traceId,
      );

      if (routerDecision.route === 'shortcut' && writeAllowed) {
        await this.handleRouterShortcut(
          routedRequest,
          routerDecision,
          conversationId,
          traceId,
          reply,
          history,
          analysis,
          emit,
        );
        return;
      }

      if (routerDecision.route === 'data') {
        await this.handleSmartDataQuery(
          routedRequest,
          analysis,
          conversationId,
          emit,
          traceId,
        );
        endSseResponse(reply);
        return;
      }

      if (routerDecision.route === 'export') {
        if (requestMode === 'ask' || requestMode === 'plan') {
          await this.emitLocalDecision(
            conversationId,
            {
              kind: 'answer',
              answer:
                `Copying matching rows to a new sheet requires **Action** mode. Switch to Action and send the same request again.`,
            },
            emit,
            { traceId, route: 'export' },
          );
          endSseResponse(reply);
          return;
        }

        const exportDecision = await this.handleFindExportQuery(
          routedRequest,
          analysis,
          conversationId,
          emit,
        );
        if (exportDecision) {
          this.logger.log(
            `Find export (router) trace=${traceId} conversation=${conversationId} mode=${activeRequest.mode ?? 'default'}`,
          );
          await this.emitLocalDecision(conversationId, exportDecision, emit, {
            traceId,
            route: 'export',
          });
          endSseResponse(reply);
          return;
        }
      }

      if (this.engine.hasOpenAi()) {
        if (routerDecision.route === 'ask') {
          // Spec 23: broad sheet overview — deterministic aggregates, skip LLM narration.
          if (isSheetOverviewRequest(routedRequest.message)) {
            emit('status', { message: 'Summarizing your sheet…' });
            const activeSheetName = this.resolveActiveSheetName(routedRequest);
            const fullData = await this.resolveActiveSheetData(
              routedRequest,
              analysis,
              activeSheetName,
              conversationId,
              emit,
            );
            const overviewAnalysis = this.sheetAnalyzer.analyze(fullData, {
              knownHeaders: analysis.headers.length ? analysis.headers : undefined,
            });
            const markdown = formatSheetOverviewMarkdown(
              buildSheetOverview(fullData, overviewAnalysis, activeSheetName),
            );
            await this.emitLocalDecision(
              conversationId,
              { kind: 'answer', answer: markdown },
              emit,
              { traceId, route: 'ask' },
            );
            endSseResponse(reply);
            return;
          }

          const ambiguityOutcome = await this.checkAmbiguity(routedRequest, analysis, history);
          if (ambiguityOutcome?.clarification) {
            await this.emitClarification(
              conversationId,
              ambiguityOutcome.clarification,
              emit,
              reply,
              traceId,
            );
            return;
          }
          if (ambiguityOutcome?.lowConfidence) {
            emit('status', {
              message: `⚠ Low confidence (${ambiguityOutcome.score}% ambiguous) — proceeding with best guess…`,
            });
          }
        }

        try {
          if (requestMode === 'plan' && routerDecision.route === 'write') {
            await this.streamPlanOnly(
              routedRequest,
              routerDecision,
              conversationId,
              traceId,
              reply,
              history,
              analysis,
              emit,
            );
          } else if (routerDecision.route === 'write' && writeAllowed) {
            await this.handleWriteRoute(
              routedRequest,
              routerDecision,
              conversationId,
              traceId,
              reply,
              history,
              analysis,
              emit,
              userId,
            );
          } else {
            await this.streamWithOpenAi(
              routedRequest,
              reply,
              conversationId,
              traceId,
              history,
              analysis,
              emit,
            );
          }
          return;
        } catch (error) {
          if (!this.shouldFallbackFromOpenAi(error)) {
            throw error;
          }
          const reason = error instanceof Error ? error.message : 'AI provider unavailable';
          localReason = `llm_fallback:${this.clipForLog(reason, 120)}`;
          // F11: keep the provider's own diagnosis instead of discarding it — the
          // local engine must not blame the user's API key for a 402/429/timeout.
          llmFailure = classifyLlmFailure(
            error instanceof LlmRequestError ? error.status : undefined,
            reason,
            true,
          );
          this.logger.warn(
            `LLM unavailable (${llmFailure.kind}), using local engine: ${reason}`,
          );
          // Task #92: persist WHY, not just that it happened — requests.log records
          // SSE events only, so the cause was previously terminal-only.
          this.workflowTrace.appendNode(traceId, {
            id: `llm_fail_${Date.now()}`,
            type: 'error',
            label: `LLM call failed (${llmFailure.kind})`,
            status: 'failed',
            meta: {
              kind: llmFailure.kind,
              status: llmFailure.status ?? null,
              detail: this.clipForLog(reason, 300),
              recoverable: true,
            },
          });
          emit('status', { message: describeLlmFailureForStatus(llmFailure) });
        }
      } else {
        llmFailure = classifyLlmFailure(undefined, undefined, false);
        emit('status', { message: describeLlmFailureForStatus(llmFailure) });
      }

      const decision = this.engine.decide(
        activeRequest.message,
        activeRequest.sheetData,
        analysis,
        history,
        resolveEngineWorkbookMeta(activeRequest),
        llmFailure,
      );
      this.logger.log(
        `AI skipped trace=${traceId} conversation=${conversationId} provider=local reason=${localReason} result=${decision.kind} durationMs=${Date.now() - startedAt}`,
      );
      await this.emitLocalDecision(conversation.conversationId, decision, emit, {
        traceId,
        route: 'local',
      });
      endSseResponse(reply);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to process your request';
      this.logger.error(message, error instanceof Error ? error.stack : undefined);
      this.workflowTrace.appendNode(traceId, {
        id: 'error',
        type: 'error',
        label: 'Error',
        status: 'failed',
        output: { message },
      });
      this.finalizeWorkflow(traceId, 'failed', {
        durationMs: Date.now() - startedAt,
        sseOutput: { error: message },
      });
      if (error instanceof WriteRouteNoActionError) {
        emit('error', {
          message: error.message,
          code: error.code,
          conversationId: error.conversationId,
        });
      } else {
        emit('error', { message });
      }
      endSseResponse(reply);
    }
  }

  /**
   * A write-route turn must never terminate as a confident prose answer with
   * zero actions. Clarifications are allowed; everything else is a bug.
   */
  private assertWriteRouteProducedActions(params: {
    conversationId: string;
    message: string;
    actionsLength: number;
    clarificationsNeeded?: string[];
  }): void {
    if (params.actionsLength > 0) {
      return;
    }

    this.logger.error('write-route turn terminated without actions', {
      conversationId: params.conversationId,
      message: params.message,
      clarificationsNeeded: params.clarificationsNeeded,
    });

    if (!params.clarificationsNeeded?.length) {
      throw new WriteRouteNoActionError(params.conversationId, params.message);
    }
  }

  private async checkAmbiguity(
    request: ConversationRequestDto,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    mongoHistory: ConversationMessageEntry[],
  ) {
    const workbookContext = resolveWorkbookContext(request, analysis, request.sheetData);
    const conversationHistory = resolveConversationHistory(request, mongoHistory);
    const quickCall = this.openRouter.isConfigured()
      ? (system: string, user: string) => this.openRouter.quickCall(system, user)
      : undefined;

    return detectAmbiguity(request.message, workbookContext, conversationHistory, quickCall);
  }

  private async emitClarification(
    conversationId: string,
    clarification: {
      question: string;
      suggestions?: string[];
      ambiguityScore: number;
    },
    emit: (event: string, data: Record<string, unknown>) => void,
    reply: FastifyReply,
    traceId?: string,
  ): Promise<void> {
    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: `[Clarification needed]: ${clarification.question}`,
      type: 'clarification',
      timestamp: new Date(),
      metadata: {
        questionOptions: clarification.suggestions,
        ambiguityScore: clarification.ambiguityScore,
      },
    });

    emit('clarification', {
      question: clarification.question,
      suggestions: clarification.suggestions,
      ambiguityScore: clarification.ambiguityScore,
    });
    emit('done', { message: 'awaiting_clarification' });
    if (traceId) {
      this.finalizeWorkflow(traceId, 'clarifying', {
        sseOutput: {
          clarification: clarification.question,
          suggestions: clarification.suggestions,
        },
      });
    }
    endSseResponse(reply);
  }

  private shouldFallbackFromOpenAi(error: unknown): boolean {
    if (error instanceof LlmRequestError) {
      return error.isRecoverable;
    }
    return false;
  }

  private resolveActiveSheetName(request: ConversationRequestDto): string {
    const richContext = request.workbookContext as { activeSheet?: string } | undefined;
    if (richContext && 'activeSheet' in richContext && richContext.activeSheet) {
      return richContext.activeSheet;
    }
    return resolveEngineWorkbookMeta(request)?.activeSheet ?? 'Sheet1';
  }

  private async collectWorkbookFindSlices(
    request: ConversationRequestDto,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    conversationId: string,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<FindExportSheetSlice[]> {
    const activeSheetName = this.resolveActiveSheetName(request);
    const richContext = resolveWorkbookContext(request, analysis, request.sheetData);
    const slices: FindExportSheetSlice[] = [];

    const activeData = await this.resolveActiveSheetData(
      request,
      analysis,
      activeSheetName,
      conversationId,
      emit,
    );
    const activeAnalysis = this.sheetAnalyzer.analyze(activeData);
    const activeMatches = this.dataQuery.collectMatches(
      request.message,
      activeData,
      activeAnalysis,
      activeSheetName,
    );
    if (activeMatches.length) {
      slices.push({
        sheetName: activeSheetName,
        sheetData: activeData,
        analysis: activeAnalysis,
        matches: activeMatches,
      });
    }

    for (const snapshot of richContext.sheets ?? []) {
      if (!snapshot?.sheetName || snapshot.sheetName === activeSheetName) continue;
      const data = await this.resolveSnapshotData(snapshot, conversationId, emit);
      if (!data.length) continue;
      const sheetAnalysis = this.sheetAnalyzer.analyze(data);
      const matches = this.dataQuery.collectMatches(
        request.message,
        data,
        sheetAnalysis,
        snapshot.sheetName,
      );
      if (!matches.length) continue;
      slices.push({
        sheetName: snapshot.sheetName,
        sheetData: data,
        analysis: sheetAnalysis,
        matches,
      });
    }

    return slices;
  }

  private async handleFindExportQuery(
    request: ConversationRequestDto,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    conversationId: string,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<EngineResponse | null> {
    const terms = this.dataQuery.extractSearchTerms(request.message);
    if (!terms.length) {
      return {
        kind: 'answer',
        answer:
          'I could not extract a search value from your message. Try: Find "Deva steels" and create a new sheet with those rows.',
      };
    }

    emit('status', { message: 'Finding matching rows and preparing export…' });

    const slices = await this.collectWorkbookFindSlices(
      request,
      analysis,
      conversationId,
      emit,
    );
    const plan = this.findExport.buildPlan(request.message, slices);
    if (!plan) return null;

    if (!plan.actions.length) {
      return {
        kind: 'answer',
        answer: plan.answer,
      };
    }

    return {
      kind: 'actions',
      answer: plan.answer,
      explanation: plan.explanation,
      actions: plan.actions,
    };
  }

  private async handleSmartDataQuery(
    request: ConversationRequestDto,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    conversationId: string,
    emit: (event: string, data: Record<string, unknown>) => void,
    traceId?: string,
  ): Promise<void> {
    emit('status', { message: 'Analyzing your sheet data…' });

    const activeSheetName = this.resolveActiveSheetName(request);
    const sheetData = await this.resolveActiveSheetData(
      request,
      analysis,
      activeSheetName,
      conversationId,
      emit,
    );
    const workbookContext = resolveWorkbookContext(request, analysis, sheetData);
    const answer = await this.smartDataQuery.handleQuery(
      request.message,
      sheetData,
      workbookContext,
      activeSheetName,
      emit,
    );

    const findPointers = this.resolveFindPointers(
      request.message,
      sheetData,
      analysis,
      activeSheetName,
    );

    await this.emitLocalDecision(
      conversationId,
      {
        kind: 'answer',
        answer,
        matches: findPointers.matches,
        selectCell: findPointers.selectCell,
      },
      emit,
      traceId ? { traceId, route: 'data' } : undefined,
    );
  }

  /** Deterministic cell targets for find/lookup so the add-in can select/pointer jump. */
  private resolveFindPointers(
    message: string,
    sheetData: unknown[][],
    _analysis: ReturnType<SheetAnalyzerService['analyze']>,
    sheetName: string,
  ): {
    matches?: ReturnType<DataQueryService['collectMatches']>;
    selectCell?: { sheetName: string; row: number; col: number };
  } {
    if (!isFindLookupMessage(message)) {
      return {};
    }

    const sheetAnalysis = this.sheetAnalyzer.analyze(sheetData);
    const matches = this.dataQuery.collectMatches(
      message,
      sheetData,
      sheetAnalysis,
      sheetName,
    );
    if (!matches.length) {
      return { matches: [] };
    }

    const first = matches[0]!;
    return {
      matches,
      selectCell: {
        sheetName: first.sheetName,
        row: first.row,
        col: first.col,
      },
    };
  }

  /** Read the active sheet's full data, fetching on demand if the payload was compressed. */
  private async resolveActiveSheetData(
    request: ConversationRequestDto,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    sheetName: string,
    conversationId: string,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<unknown[][]> {
    let sheetData = request.sheetData;
    const declaredRows = this.resolveDeclaredRowCount(request);
    const isTruncated = Boolean(request.sheetCompression?.truncated);
    const canFetch = Boolean(request.sheetCompression?.onDemandFetchEnabled);
    const dataIncomplete = declaredRows > sheetData.length;

    if ((isTruncated || dataIncomplete) && canFetch) {
      const lastCol = analysis.columnLetters[Math.max(0, analysis.columnCount - 1)] ?? 'A';
      const range = `A1:${lastCol}${declaredRows}`;
      try {
        const fetched = await this.toolBridge.waitForRangeData(
          conversationId,
          { name: 'get_range_data', sheet: sheetName, range },
          emit,
        );
        if (fetched.values?.length) {
          sheetData = fetched.values;
          this.logger.log(`Find query fetched ${sheetData.length} rows from ${sheetName}!${range}`);
        } else if (fetched.error) {
          this.logger.warn(`Find query range fetch error: ${fetched.error}`);
        }
      } catch (error) {
        const reason = error instanceof Error ? error.message : 'range fetch failed';
        this.logger.warn(`Find query range fetch failed: ${reason}`);
      }
    }

    return sheetData;
  }

  /** Read a non-active sheet's data from its snapshot, fetching the full range on demand. */
  private async resolveSnapshotData(
    snapshot: SheetSnapshot,
    conversationId: string,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<unknown[][]> {
    const sampled = (snapshot.sampleData ?? []) as unknown[][];
    const meta = snapshot.compressionMeta;
    const truncated = Boolean(meta?.truncated);
    const canFetch = Boolean(meta?.onDemandFetchEnabled);

    if (truncated && canFetch && snapshot.usedRange) {
      try {
        const fetched = await this.toolBridge.waitForRangeData(
          conversationId,
          { name: 'get_range_data', sheet: snapshot.sheetName, range: snapshot.usedRange },
          emit,
        );
        if (fetched.values?.length) {
          this.logger.log(
            `Find query fetched ${fetched.values.length} rows from ${snapshot.sheetName}!${snapshot.usedRange}`,
          );
          return fetched.values;
        }
        if (fetched.error) {
          this.logger.warn(`Find cross-sheet fetch error (${snapshot.sheetName}): ${fetched.error}`);
        }
      } catch (error) {
        const reason = error instanceof Error ? error.message : 'range fetch failed';
        this.logger.warn(`Find cross-sheet fetch failed (${snapshot.sheetName}): ${reason}`);
      }
    }

    return sampled;
  }

  private async emitLocalDecision(
    conversationId: string,
    decision: EngineResponse,
    emit: (event: string, data: Record<string, unknown>) => void,
    workflow?: {
      traceId: string;
      route?: string;
      tier?: number;
      changeSetId?: string;
      status?: WorkflowTraceStatus;
    },
  ): Promise<void> {
    if (decision.kind === 'question') {
      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: decision.question,
        type: 'question',
        timestamp: new Date(),
        metadata: {
          questionOptions: decision.options,
          pendingIntent: decision.pendingIntent,
        },
      });

      emit('question', {
        question: decision.question,
        options: decision.options,
      });
      if (workflow?.traceId) {
        this.finalizeWorkflow(workflow.traceId, 'clarifying', {
          route: workflow.route,
          tier: workflow.tier,
          sseOutput: { question: decision.question, options: decision.options },
        });
      }
      return;
    }

    if (decision.kind === 'actions') {
      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: decision.answer,
        type: 'answer',
        timestamp: new Date(),
        metadata: this.buildWriteMetadata(decision.actions, workflow?.changeSetId),
      });

      emit('answer', { answer: decision.answer });
      emit('actions', {
        actions: decision.actions,
        explanation: decision.explanation,
      });
      emit('conversation_end', { summary: 'Changes applied.' });
      await this.markCompleted(conversationId);
      if (workflow?.traceId) {
        this.finalizeWorkflow(workflow.traceId, workflow.status ?? 'completed', {
          route: workflow.route,
          tier: workflow.tier,
          changeSetId: workflow.changeSetId,
          sseOutput: {
            kind: 'actions',
            answer: decision.answer,
            actionTypes: decision.actions.map((a) => a.type),
            explanation: decision.explanation,
          },
        });
      }
      return;
    }

    const answerMetadata = await this.buildAnswerPersistMetadata(
      conversationId,
      decision.answer,
      false,
    );
    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: decision.answer,
      type: 'answer',
      timestamp: new Date(),
      ...(answerMetadata ? { metadata: answerMetadata } : {}),
    });

    const matches = decision.kind === 'answer' ? decision.matches : undefined;
    const selectCell =
      decision.kind === 'answer'
        ? decision.selectCell ??
          (matches?.[0]
            ? {
                sheetName: matches[0].sheetName,
                row: matches[0].row,
                col: matches[0].col,
              }
            : undefined)
        : undefined;

    emit('answer', {
      answer: decision.answer,
      matches,
    });
    if (matches?.length) {
      emit('matches', { matches, summary: decision.answer });
    }
    if (selectCell) {
      emit('select_cell', selectCell);
    }
    emit('conversation_end', { summary: 'Ready for your next message.' });
    await this.markCompleted(conversationId);
    if (workflow?.traceId) {
      this.finalizeWorkflow(workflow.traceId, workflow.status ?? 'completed', {
        route: workflow.route,
        tier: workflow.tier,
        sseOutput: {
          kind: 'answer',
          answer: decision.answer,
          matchCount: matches?.length ?? 0,
          pendingWritePlan: Boolean(answerMetadata?.pendingWritePlan),
        },
      });
    }
  }

  /**
   * Explicit tier dispatch for route=write requests (Tier 0–3).
   * Tier 3 delegates to streamWithOrchestrator() unchanged.
   */
  private async handleWriteRoute(
    routedRequest: ConversationRequestDto,
    routerDecision: RouterDecision,
    conversationId: string,
    traceId: string,
    reply: FastifyReply,
    history: ConversationMessageEntry[],
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    emit: (event: string, data: Record<string, unknown>) => void,
    userId?: string,
  ): Promise<void> {
    const startedAt = Date.now();
    const outcome = this.resolveInitialWriteOutcome(routerDecision);

    try {
      const richWorkbookContext = resolveWorkbookContext(
        routedRequest,
        analysis,
        routedRequest.sheetData,
      );
      const deleteActions = tryLocalDeleteSheetActions(
        routedRequest.message,
        richWorkbookContext,
      );
      if (deleteActions?.length) {
        const sheetNames = deleteActions
          .map((action) => action.sheetName)
          .filter(Boolean) as string[];
        outcome.llmCallCount = 0;
        this.logger.log(
          `Delete sheet (deterministic) trace=${traceId} conversation=${conversationId} sheets=${sheetNames.join(',')}`,
        );
        await this.emitLocalDecision(
          conversationId,
          {
            kind: 'actions',
            answer: buildDeleteSheetAnswer(sheetNames),
            explanation: 'Removed the requested worksheet tab(s).',
            actions: deleteActions,
          },
          emit,
          { traceId, route: 'write', tier: 0 },
        );
        endSseResponse(reply);
        return;
      }

      const classifiedTier = (routerDecision.complexity ?? 3) as 0 | 1 | 2 | 3;
      const tieringMode = getComplexityTieringMode();
      const complexity = resolveExecutableTier(classifiedTier, tieringMode);
      const actionHint = routerDecision.actionHint;
      this.logWorkflowTier(traceId, complexity, actionHint, {
        classifiedTier,
        tieringMode,
      });
      const agentContext = buildAgentWorkbookContext(
        richWorkbookContext,
        routedRequest.sheetData,
        analysis,
      );

      if (complexity <= 1) {
        if (complexity === 0 && actionHint) {
          const tier0Result = this.tier0Direct.resolve(
            actionHint,
            routedRequest.message,
            agentContext,
          );
          if (tier0Result) {
            outcome.tier = 0;
            outcome.llmCallCount = 0;
            this.logger.log(
              `Tier 0 direct trace=${traceId} conversation=${conversationId} actionHint=${actionHint} actions=${tier0Result.actions.map((a) => a.type).join(',')}`,
            );
            await this.streamTier0Result(
              conversationId,
              traceId,
              routedRequest.message,
              tier0Result,
              agentContext,
              routerDecision.assumption,
              emit,
            );
            endSseResponse(reply);
            return;
          }

          this.logger.warn(
            `[${traceId}] Tier 0 downgrade reason=implicit_target actionHint=${actionHint}`,
          );
        }

        if (actionHint) {
          try {
            const tier1Result = await this.tier1SingleAction.execute(
              routedRequest.message,
              actionHint,
              agentContext,
            );
            // TASKS.md #165 — the word-based lane guess was made before any work
            // existed; now that it does, check it. Escalating discards this
            // lane's LLM call, which is why the thresholds are conservative.
            const t1Escalation = assessTierEscalation(tier1Result.actions);
            if (t1Escalation.escalate) {
              this.logger.warn(
                `[${traceId}] Tier 1 escalating to planner: ${t1Escalation.reason}`,
              );
              outcome.escalatedFrom = 1;
              outcome.escalationReason = t1Escalation.reason ?? undefined;
            } else if (tier1Result.actions.length > 0) {
              outcome.finalActionCount = tier1Result.actions.length;
              outcome.tier = 1;
              outcome.llmCallCount = 1;
              this.logger.log(
                `Tier 1 single-action trace=${traceId} conversation=${conversationId} actionHint=${actionHint} action=${tier1Result.actions[0].type}`,
              );
              await this.streamTier1Result(
                conversationId,
                traceId,
                routedRequest.message,
                tier1Result,
                actionHint,
                agentContext,
                routerDecision.assumption,
                emit,
              );
              endSseResponse(reply);
              return;
            }
          } catch (error) {
            if (
              error instanceof Error &&
              error.message === 'numeric_find_replace_escalation_required'
            ) {
              this.logger.warn(
                `[${traceId}] Tier 1 blocked numeric FIND_REPLACE — falling through to orchestrator`,
              );
            } else {
              throw error;
            }
          }
        }
      } else if (complexity === 2 && actionHint) {
        // CREDIT_SYSTEM.md CD-4 — pre-flight gate before the LLM call, not after.
        // userId is only absent for the eval-bypass auth path (auth.guard.ts) or
        // a session with no user id; skip the gate rather than block a caller
        // this codebase doesn't yet have an identity to charge.
        if (userId) {
          const gateResult = await this.creditGate.checkBalance(userId, 'FORMULA_GENERATE_OR_FIX');
          if (!gateResult.allowed && gateResult.reason === 'insufficient_balance') {
            emit('error', {
              message: 'You are out of credits for this action.',
              code: 'INSUFFICIENT_CREDIT',
              availableBalance: gateResult.availableBalance,
              requiredCredits: gateResult.requiredCredits,
            });
            await this.markCompleted(conversationId);
            this.finalizeWorkflow(traceId, 'failed', {
              route: 'write',
              tier: 2,
              sseOutput: { error: 'insufficient_credit' },
            });
            endSseResponse(reply);
            return;
          }
        }

        const basePrompt =
          routedRequest.promptContext ?? richWorkbookContext.prompt_context ?? undefined;
        const { enrichedContext } = this.enrichAgentContext(
          agentContext,
          basePrompt,
          history,
          routedRequest.message,
        );
        const tier2Result = await this.tier2GenerateVerify.execute(
          routedRequest.message,
          actionHint,
          enrichedContext,
          traceId,
          { conversationId, toolEmit: emit },
        );
        // TASKS.md #165 — same check as Tier 1. Tier 2 is where this matters
        // most: it can legitimately emit a handful of actions, so a result that
        // creates several sheets is a build that slipped past the classifier.
        const t2Escalation = assessTierEscalation(tier2Result.actions);
        if (t2Escalation.escalate) {
          this.logger.warn(
            `[${traceId}] Tier 2 escalating to planner: ${t2Escalation.reason}`,
          );
          outcome.escalatedFrom = 2;
          outcome.escalationReason = t2Escalation.reason ?? undefined;
        } else {
        outcome.finalActionCount = tier2Result.actions.length;
        outcome.tier = 2;
        // Executor+Verifier, plus optional Bug 1 retry (+ verify) and Bug 4 tool follow-up.
        outcome.llmCallCount = tier2Result.toolFollowUp
          ? 5
          : tier2Result.retried
            ? 4
            : 2;
        await this.streamTier2Result(
          routedRequest,
          conversationId,
          traceId,
          analysis,
          richWorkbookContext,
          enrichedContext,
          tier2Result,
          reply,
          emit,
          routerDecision.assumption,
          userId,
        );
        return;
        }
      }

      outcome.tier = 3;
      outcome.llmCallCount = 3;
      // Over-sorting is only visible if lane 3's OWN action count is recorded:
      // a three-minute pipeline that produced two actions was the wrong lane.
      const reportActionCount = (count: number) => {
        outcome.finalActionCount = count;
      };
      await this.streamWithOrchestrator(
        {
          ...routedRequest,
          message: stripSheetMentions(routedRequest.message),
        },
        reply,
        conversationId,
        traceId,
        history,
        analysis,
        emit,
        routerDecision.assumption,
        (routerDecision.complexity ?? 3) as 0 | 1 | 2 | 3,
        reportActionCount,
        userId,
      );
    } finally {
      const classifiedTier = (routerDecision.complexity ?? 3) as 0 | 1 | 2 | 3;
      const tieringMode = getComplexityTieringMode();
      this.structuredLogger.logTierDecision({
        traceId,
        message: routedRequest.message,
        tier: outcome.tier,
        classifiedTier,
        tieringMode,
        shadowed: classifiedTier !== outcome.tier,
        matchedBy: routerDecision.matchedBy ?? 'llm-fallback',
        actionHint: routerDecision.actionHint ?? '',
        llmCallCount: outcome.llmCallCount,
        durationMs: Date.now() - startedAt,
        // TASKS.md #165 — the three fields that make mis-sorting MEASURABLE
        // rather than a matter of opinion. Pairing the lane with what it
        // actually produced is what turns "is the classifier any good?" into a
        // query: a lane 1/2 run with a large `finalActionCount` was
        // under-sorted; a lane 3 run with two or three actions was over-sorted.
        escalatedFrom: outcome.escalatedFrom,
        escalationReason: outcome.escalationReason,
        finalActionCount: outcome.finalActionCount,
      });
      if (outcome.escalatedFrom) {
        this.logger.warn(
          `[${traceId}] ROUTING MISS: classified tier ${outcome.escalatedFrom}, escalated to ${outcome.tier} — ${outcome.escalationReason}`,
        );
      }
    }
  }

  private resolveInitialWriteOutcome(routerDecision: RouterDecision): {
    tier: 0 | 1 | 2 | 3;
    llmCallCount: number;
    /** Set when a fast lane bailed upward mid-flight. TASKS.md #165. */
    escalatedFrom?: 1 | 2;
    escalationReason?: string;
    /** Actions the run finally produced — pairs with `tier` to expose mis-sorting. */
    finalActionCount?: number;
  } {
    const complexity = routerDecision.complexity ?? 3;
    if (complexity === 0) {
      return { tier: 0, llmCallCount: 0 };
    }
    if (complexity === 1) {
      return { tier: 1, llmCallCount: 1 };
    }
    if (complexity === 2) {
      return { tier: 2, llmCallCount: 2 };
    }
    return { tier: 3, llmCallCount: 3 };
  }

  /**
   * Plan mode for write routes: describe or generate proposals without ChangeSet / apply.
   */
  private async streamPlanOnly(
    routedRequest: ConversationRequestDto,
    routerDecision: RouterDecision,
    conversationId: string,
    traceId: string,
    reply: FastifyReply,
    history: ConversationMessageEntry[],
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<void> {
    const startedAt = Date.now();
    const complexity = routerDecision.complexity ?? 3;
    const actionHint = routerDecision.actionHint ?? '';
    let llmCallCount = 0;

    emit('status', { message: 'Building a plan without applying changes…' });

    try {
      const richWorkbookContext = resolveWorkbookContext(
        routedRequest,
        analysis,
        routedRequest.sheetData,
      );
      const agentContext = buildAgentWorkbookContext(
        richWorkbookContext,
        routedRequest.sheetData,
        analysis,
      );
      const basePrompt =
        routedRequest.promptContext ?? richWorkbookContext.prompt_context ?? undefined;
      const { enrichedContext, promptContext } = this.enrichAgentContext(
        agentContext,
        basePrompt,
        history,
        routedRequest.message,
      );
      const conversationHistory = resolveConversationHistory(routedRequest, history).map(
        (entry) => ({
          role: entry.role as 'user' | 'assistant',
          content: entry.content,
        }),
      );

      if (complexity <= 1) {
        const description = this.describeIntendedAction(
          routedRequest.message,
          routerDecision,
          agentContext,
        );
        await this.emitPlanOnly({
          conversationId,
          prompt: routedRequest.message,
          summary: 'Single-step action preview',
          steps: [{ title: description }],
          tier: complexity === 0 ? 0 : 1,
          answer: `Here's what would happen in Action mode:\n\n${description}`,
          emit,
        });
        endSseResponse(reply);
        return;
      }

      if (complexity === 2 && actionHint) {
        llmCallCount = 1;
        emit('thinking', { message: '🔍 Generating a proposed change (no verification yet)…' });
        const generateResult = await this.tier2GenerateVerify.generateOnly(
          routedRequest.message,
          actionHint,
          enrichedContext,
          traceId,
        );
        const steps =
          generateResult.actions.length > 0
            ? generateResult.actions.map((action) => ({
                title: this.describeProposedSheetAction(action),
                detail: actionHint,
              }))
            : [{ title: generateResult.answer }];

        await this.emitPlanOnly({
          conversationId,
          prompt: routedRequest.message,
          summary: `Proposed ${actionHint.replace(/_/g, ' ').toLowerCase()}`,
          steps,
          proposedActions: generateResult.actions,
          tier: 2,
          answer: generateResult.answer,
          emit,
        });
        endSseResponse(reply);
        return;
      }

      llmCallCount = 1;
      emit('thinking', { message: '🧠 Building a step-by-step plan across your workbook…' });

      const plan = await this.orchestrator.planOnly({
        prompt: routedRequest.message,
        context: enrichedContext,
        conversationHistory,
        promptContext,
        conversationId,
        correlationId: traceId,
        complexity: (complexity ?? 3) as 0 | 1 | 2 | 3,
      });

      if (plan.clarificationsNeeded.length > 0) {
        const question = plan.clarificationsNeeded.join(' ');
        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: `[Clarification needed]: ${question}`,
          type: 'clarification',
          timestamp: new Date(),
        });
        emit('clarification', { question, suggestions: [], ambiguityScore: 0 });
        emit('done', { message: 'awaiting_clarification' });
        endSseResponse(reply);
        return;
      }

      const planPayload = this.buildPlanPayload(plan, routedRequest.message, enrichedContext);
      const answer =
        planPayload.steps.length > 0
          ? `Here's a ${planPayload.steps.length}-step plan. Review it, then run it as an action to preview and apply the changes.`
          : 'I could not break this request into concrete steps. Try rephrasing, or switch to Action mode.';

      await this.emitPlanOnly({
        conversationId,
        prompt: routedRequest.message,
        summary: planPayload.summary,
        steps: planPayload.steps,
        affectedSheets: planPayload.affectedSheets,
        estimatedRows: planPayload.estimatedRows,
        safestApproach: planPayload.safestApproach,
        tier: 3,
        answer,
        emit,
      });
      endSseResponse(reply);
    } finally {
      this.structuredLogger.logTierDecision({
        traceId,
        message: routedRequest.message,
        tier: (complexity <= 1 ? complexity : complexity === 2 ? 2 : 3) as 0 | 1 | 2 | 3,
        matchedBy: routerDecision.matchedBy ?? 'llm-fallback',
        actionHint: routerDecision.actionHint ?? '',
        llmCallCount,
        durationMs: Date.now() - startedAt,
      });
    }
  }

  private describeIntendedAction(
    message: string,
    routerDecision: RouterDecision,
    agentContext: AgentWorkbookContext,
  ): string {
    const actionHint = routerDecision.actionHint ?? '';
    if (routerDecision.complexity === 0 && actionHint) {
      const tier0Result = this.tier0Direct.resolve(actionHint, message, agentContext);
      if (tier0Result) {
        const summary = this.buildTier0Answer(tier0Result.actions);
        return `${summary} (${actionHint.replace(/_/g, ' ').toLowerCase()})`;
      }
    }
    return this.buildActionHintDescription(message, actionHint);
  }

  private buildActionHintDescription(message: string, actionHint: string): string {
    const labels: Record<string, string> = {
      CELL_FORMAT: 'Apply formatting to the specified cells',
      FREEZE_PANES: 'Freeze the top row or panes on the active sheet',
      VISIBILITY_TOGGLE: 'Change row, column, or sheet visibility',
      ROW_COL_STRUCTURE: 'Insert or delete rows or columns',
      SORT_OR_FILTER: 'Sort or filter data based on your criteria',
      FIND_REPLACE: 'Find and replace matching values',
      CONDITIONAL_FORMAT: 'Apply conditional formatting rules',
      HEADER_FORMAT: 'Format the header row',
      COPY_FILL: 'Copy formatting or fill values down a column',
      FORMULA_GEN: 'Generate a formula for the requested calculation',
      PIVOT_TABLE: 'Create or update a pivot table',
      CHART: 'Create or update a chart',
      DUPLICATE_CHECK: 'Identify duplicate values',
      DATA_VALIDATION: 'Add data validation or dropdown rules',
      ERROR_FIX: 'Fix formula errors in the affected cells',
    };
    const label =
      labels[actionHint] ??
      (actionHint
        ? `Perform a ${actionHint.replace(/_/g, ' ').toLowerCase()} operation`
        : 'Apply a single change to your workbook');
    return `${label}: "${this.clipForLog(message, 200)}"`;
  }

  private describeProposedSheetAction(action: SheetAction): string {
    if (action.type === 'SET_CELL' && action.row !== undefined && action.col !== undefined) {
      const col = String.fromCharCode(65 + action.col);
      return `Set cell ${col}${action.row + 1} to ${action.value ?? ''}`;
    }
    if (action.type === 'SET_FORMULA' && action.row !== undefined && action.col !== undefined) {
      const col = String.fromCharCode(65 + action.col);
      return `Set formula in ${col}${action.row + 1}${action.formula ? `: ${action.formula}` : ''}`;
    }
    if (action.type === 'FORMAT_RANGE') {
      return 'Apply formatting to the target range';
    }
    return action.type.replace(/_/g, ' ').toLowerCase();
  }

  private async emitPlanOnly(params: {
    conversationId: string;
    prompt: string;
    summary?: string;
    steps: { title: string; detail?: string }[];
    proposedActions?: SheetAction[];
    affectedSheets?: string[];
    estimatedRows?: number;
    safestApproach?: string;
    tier: 0 | 1 | 2 | 3;
    answer: string;
    emit: (event: string, data: Record<string, unknown>) => void;
  }): Promise<void> {
    await this.saveMessage(params.conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: params.answer,
      type: 'answer',
      timestamp: new Date(),
      metadata: params.proposedActions?.length
        ? { actions: params.proposedActions }
        : undefined,
    });

    params.emit('answer', { answer: params.answer, tier: params.tier });
    params.emit('plan_only', {
      prompt: params.prompt,
      summary: params.summary,
      steps: params.steps,
      proposedActions: params.proposedActions,
      affectedSheets: params.affectedSheets ?? [],
      estimatedRows: params.estimatedRows,
      safestApproach: params.safestApproach,
      tier: params.tier,
    });
    params.emit('conversation_end', {
      summary: 'Plan ready — run as action to apply.',
      tier: params.tier,
    });
    await this.markCompleted(params.conversationId);
  }

  private async streamTier0Result(
    conversationId: string,
    traceId: string,
    message: string,
    result: Tier0Result,
    agentContext: AgentWorkbookContext,
    assumption: string | undefined,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<void> {
    const answer = this.buildTier0Answer(result.actions);
    await this.emitTierActions({
      conversationId,
      tier: 0,
      answer,
      processingLabel: tierProcessingLabel(0),
      actions: result.actions,
      emit,
      traceId,
      prompt: message,
      agentContext,
      assumption,
    });
    this.logger.debug(`[${traceId}] Tier 0 completed message="${this.clipForLog(message, 120)}"`);
  }

  private async streamTier1Result(
    conversationId: string,
    traceId: string,
    message: string,
    result: { actions: SheetAction[]; answer: string; model?: string },
    actionHint: string,
    agentContext: AgentWorkbookContext,
    assumption: string | undefined,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<void> {
    await this.emitTierActions({
      conversationId,
      tier: 1,
      answer: result.answer,
      processingLabel: tierProcessingLabel(1, actionHint),
      actions: result.actions,
      emit,
      traceId,
      prompt: message,
      agentContext,
      assumption,
      model: result.model,
    });
  }

  private async streamTier2Result(
    request: ConversationRequestDto,
    conversationId: string,
    traceId: string,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    richWorkbookContext: ReturnType<typeof resolveWorkbookContext>,
    agentContext: AgentWorkbookContext,
    result: Awaited<ReturnType<Tier2GenerateVerifyService['execute']>>,
    reply: FastifyReply,
    emit: (event: string, data: Record<string, unknown>) => void,
    assumption?: string,
    userId?: string,
  ): Promise<void> {
    if (result.actions.length === 0) {
      this.assertWriteRouteProducedActions({
        conversationId,
        message: request.message,
        actionsLength: 0,
      });
    }

    if (!result.verifierPassed) {
      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: result.answer,
        type: 'answer',
        timestamp: new Date(),
      });
      emit('answer', { answer: result.answer, tier: 2 });
      emit('conversation_end', {
        summary: 'Verification failed.',
        tier: 2,
      });
      await this.markCompleted(conversationId);
      this.finalizeWorkflow(traceId, 'failed', {
        route: 'write',
        tier: 2,
        durationMs: result.durationMs,
        sseOutput: { answer: result.answer, verifierPassed: false },
      });
      endSseResponse(reply);
      return;
    }

    const actions = this.engine.finalizeActions(
      result.actions,
      analysis,
      richWorkbookContext,
      request.message,
      agentContext.priorTurnActions,
    );

    this.assertWriteRouteProducedActions({
      conversationId,
      message: request.message,
      actionsLength: actions.length,
    });

    const changeSet = await this.changeSetService.createPreview({
      conversationId,
      traceId,
      prompt: request.message,
      context: agentContext,
      actions,
      provenance: {
        sourceRefs: result.sourceRefs,
        workbookId: agentContext.activeSheetName || 'workbook',
        activeSheetName: agentContext.activeSheetName,
      },
    });

    const processingLabel = tierProcessingLabel(2);
    const userFacingSummary = buildUserFacingSummary({
      answer: result.answer,
      actions,
      changes: changeSet.changes,
      assumption,
      activeSheetName: agentContext.activeSheetName,
    });
    const internalDetails = buildInternalDetails({
      tier: 2,
      processingLabel,
      assumption,
      actions,
      legacyExplanation: processingLabel,
    });

    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: result.answer,
      type: 'answer',
      timestamp: new Date(),
      metadata: this.buildWriteMetadata(actions, changeSet.changeSetId),
    });

    emit('answer', { answer: result.answer, tier: 2 });
    emit('actions', {
      actions,
      explanation: processingLabel,
      userFacingSummary,
      internalDetails,
      changeSetId: changeSet.changeSetId,
      changes: changeSet.changes,
      irreversibleActionTypes: changeSet.irreversibleActionTypes,
      tier: 2,
      durationMs: result.durationMs,
    });
    emit('conversation_end', { summary: 'Review changes and accept or reject.', tier: 2 });
    await this.markCompleted(conversationId);
    if (userId) {
      // CREDIT_SYSTEM.md CD-3 — debit once, at turn completion, now that a real
      // ChangeSet exists. Rejection later is a workbook-state decision, not a
      // billing one (CD-3) — this fires regardless of whether the user accepts.
      const debitResult = await this.creditLedger.debit(userId, 'FORMULA_GENERATE_OR_FIX', 1, {
        conversationId,
        changeSetId: changeSet.changeSetId,
      });
      if (debitResult.debited && debitResult.balances) {
        emit('credits', {
          planCredits: debitResult.balances.planCredits,
          purchasedCredits: debitResult.balances.purchasedCredits,
          oneTimeCredits: debitResult.balances.oneTimeCredits,
          debited: resolveCreditCost('FORMULA_GENERATE_OR_FIX'),
          actionType: 'FORMULA_GENERATE_OR_FIX',
        });
      }
      // debitResult.debited === false here means the balance was consumed by a
      // race since the pre-flight gate check (CD-6) — CD-4 treats this as an
      // accepted, rare timing artifact, not a reason to fail an already-verified
      // turn or withhold the ChangeSet the user is about to see.
    }
    this.logWorkflowChangeSet(traceId, changeSet.changeSetId, actions, changeSet.changes.length);
    this.finalizeWorkflow(traceId, 'awaiting_accept', {
      changeSetId: changeSet.changeSetId,
      route: 'write',
      tier: 2,
      durationMs: result.durationMs,
      sseOutput: {
        answer: result.answer,
        changeSetId: changeSet.changeSetId,
        actionTypes: actions.map((a) => a.type),
        changesLength: changeSet.changes.length,
        userFacingSummary,
      },
    });
    endSseResponse(reply);
  }

  private buildTier0Answer(actions: SheetAction[]): string {
    const first = actions[0];
    if (!first) return 'Done.';
    if (first.type === 'FORMAT_RANGE') {
      const format = first.format;
      if (format?.bold) return 'Applied bold formatting to the requested cells.';
      if (format?.italic) return 'Applied italic formatting to the requested cells.';
      if (format?.underline) return 'Applied underline formatting to the requested cells.';
      return 'Applied formatting to the requested cells.';
    }
    return buildShortcutAnswer(actions);
  }

  private async emitTierActions(params: {
    conversationId: string;
    tier: 0 | 1;
    answer: string;
    processingLabel: string;
    actions: SheetAction[];
    emit: (event: string, data: Record<string, unknown>) => void;
    traceId: string;
    prompt: string;
    agentContext: AgentWorkbookContext;
    assumption?: string;
    model?: string;
  }): Promise<void> {
    const {
      conversationId,
      tier,
      processingLabel,
      emit,
      traceId,
      prompt,
      agentContext,
      assumption,
      model,
    } = params;

    const answer = sanitizeAnswerForUser(params.answer);

    const actions = annotateExplicitOverwriteConfirmation(
      params.actions,
      prompt,
      agentContext.priorTurnActions ?? [],
    );

    const changeSet = await this.changeSetService.createPreview({
      conversationId,
      traceId,
      prompt,
      context: agentContext,
      actions,
      provenance: {
        sourceRefs: buildWorkbookSourceRefsFromActions(
          actions,
          agentContext.activeSheetName || 'workbook',
          agentContext.activeSheetName,
        ),
        workbookId: agentContext.activeSheetName || 'workbook',
        activeSheetName: agentContext.activeSheetName,
      },
    });

    const userFacingSummary = buildUserFacingSummary({
      answer,
      actions,
      changes: changeSet.changes,
      assumption,
      activeSheetName: agentContext.activeSheetName,
    });
    const internalDetails = buildInternalDetails({
      tier,
      model,
      processingLabel,
      assumption,
      actions,
      legacyExplanation: processingLabel,
    });

    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: answer,
      type: 'answer',
      timestamp: new Date(),
      metadata: this.buildWriteMetadata(actions, changeSet.changeSetId),
    });

    emit('answer', { answer, tier });
    emit('actions', {
      actions,
      explanation: processingLabel,
      userFacingSummary,
      internalDetails,
      changeSetId: changeSet.changeSetId,
      changes: changeSet.changes,
      irreversibleActionTypes: changeSet.irreversibleActionTypes,
      tier,
    });
    emit('conversation_end', { summary: 'Review changes and accept or reject.', tier });
    await this.markCompleted(conversationId);
    this.logWorkflowChangeSet(traceId, changeSet.changeSetId, actions, changeSet.changes.length);
    this.finalizeWorkflow(traceId, 'awaiting_accept', {
      changeSetId: changeSet.changeSetId,
      route: 'write',
      tier,
      sseOutput: {
        answer,
        changeSetId: changeSet.changeSetId,
        actionTypes: actions.map((a) => a.type),
        changesLength: changeSet.changes.length,
        userFacingSummary,
      },
    });
  }

  /**
   * Plans a Tier 3 build, and — when the plan has more than one dependency wave
   * — persists it as an `agent_run`, executes ONLY wave 0, emits its Accept
   * card, and ends the stream (STEPWISE_EXECUTION.md SD-1/SD-3).
   *
   * Returns true when it took ownership of the response. Returns false to mean
   * "not stepwise after all" — a single-wave plan, or a plan that must ask a
   * clarification first — in which case the caller runs the unchanged one-shot
   * path. Deliberately never throws for a can't-do-stepwise reason: falling
   * back to the behaviour that already works beats failing a real request.
   */
  private async tryStartStepwiseRun(opts: {
    request: ConversationRequestDto;
    reply: FastifyReply;
    conversationId: string;
    traceId: string;
    emit: (event: string, data: Record<string, unknown>) => void;
    sseEmitter: SseEmitter;
    enrichedContext: AgentWorkbookContext;
    promptContext: string;
    conversationHistory: { role: 'user' | 'assistant'; content: string }[];
    routerAssumption?: string;
    complexity?: 0 | 1 | 2 | 3;
    telemetry: LlmCallTelemetry;
    userId?: string;
    startedAt: number;
  }): Promise<boolean> {
    const { plan, openQuestions, mustAsk } = await this.orchestrator.planForStepwiseRun(
      {
        prompt: opts.request.message,
        context: opts.enrichedContext,
        conversationHistory: opts.conversationHistory,
        promptContext: opts.promptContext,
        conversationId: opts.conversationId,
        correlationId: opts.traceId,
        toolEmit: opts.emit,
        routerAssumption: opts.routerAssumption,
        complexity: opts.complexity ?? 3,
      },
      opts.sseEmitter,
      opts.telemetry,
    );

    if (mustAsk) {
      await this.saveMessage(opts.conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: '[Clarification needed]',
        type: 'clarification',
        timestamp: new Date(),
      });
      opts.emit('done', { message: 'awaiting_clarification' });
      endSseResponse(opts.reply);
      return true;
    }

    const waves = computeExecutionWaves(plan.subtasks);
    if (!shouldRunStepwise(waves.length)) {
      // Single-wave plan: gating it would add a round trip and buy nothing.
      // The one-shot path re-plans, which costs a second Planner call — an
      // accepted, bounded cost for keeping the two paths from sharing mutable
      // plan state across a fallback boundary.
      this.logger.log(
        `Stepwise declined trace=${opts.traceId} waves=${waves.length} — falling back to one-shot`,
      );
      return false;
    }

    const run = await this.agentRunState.createRun({
      conversationId: opts.conversationId,
      userId: opts.userId,
      traceId: opts.traceId,
      prompt: opts.request.message,
      subtasks: plan.subtasks,
      context: opts.enrichedContext,
      promptContext: opts.promptContext,
      conversationHistory: opts.conversationHistory,
      routerAssumption: opts.routerAssumption,
    });

    this.logger.log(
      `Stepwise run ${run.runId} started trace=${opts.traceId} subtasks=${plan.subtasks.length} waves=${waves.length}`,
    );

    // STEPWISE_EXECUTION.md §3 originally called for emitting the whole plan
    // up front via the existing 'plan' SSE event, so the user sees the shape
    // of the build before approving its first step. Reverted: 'plan' is the
    // Plan MODE contract — a read-only preview whose card says "review it,
    // then run it as an action" and renders a "Run as Action" button that
    // re-sends the prompt. Emitting it here, while the build is ALREADY
    // executing, produced exactly that confusing button on a live run (a real
    // user report). The per-wave "Step N of M" Accept cards already show
    // progress as it happens; a proper plan-overview needs its own event type
    // and rendering (no run button, informational only) — scoped as separate
    // follow-up work rather than bolted on here under time pressure.
    if (openQuestions.length > 0) {
      opts.emit('status', {
        message: `Proceeding under an assumption — ${openQuestions[0]}`,
      });
    }

    await this.executeStepwiseWave(run, opts.reply, opts.emit, opts.sseEmitter, opts.telemetry);
    return true;
  }

  /**
   * Runs the run's next executable wave, emits its Accept card, and ends the
   * stream with `wave_ready` — the event that distinguishes "paused, call
   * /continue" from "finished" (STEPWISE_EXECUTION.md §3).
   */
  private async executeStepwiseWave(
    run: AgentRunDocument,
    reply: FastifyReply,
    emit: (event: string, data: Record<string, unknown>) => void,
    sseEmitter: SseEmitter,
    telemetry: LlmCallTelemetry,
  ): Promise<void> {
    const next = this.agentRunState.nextExecutableWave(run);

    if (!next) {
      await this.finishStepwiseRun(run, reply, emit);
      return;
    }

    await this.agentRunState.markStatus(run, 'running');

    // Everything earlier waves produced, so this wave's Executor and shadow
    // workbook see the sheets those waves created (SD-1).
    const priorActions = run.subtaskStates
      .filter((state) => state.actions.length > 0)
      .map((state) => ({
        subtask: run.subtasks.find((subtask) => subtask.id === state.subtaskId)!,
        actions: state.actions as SheetAction[],
      }))
      .filter((entry) => Boolean(entry.subtask));

    const waveResult = await this.orchestrator.runStepwiseWave(
      {
        prompt: run.prompt,
        context: run.context as AgentWorkbookContext,
        conversationHistory: run.conversationHistory,
        promptContext: run.promptContext,
        conversationId: run.conversationId,
        correlationId: run.traceId,
        toolEmit: emit,
        routerAssumption: run.routerAssumption,
        complexity: 3,
        waveSubtasks: next.subtasks,
        priorActions,
      },
      sseEmitter,
      telemetry,
    );

    await this.agentRunState.recordWaveResult(
      run,
      next.waveIndex,
      next.subtasks.map((subtask) => {
        const completed = waveResult.completedSubtasks.find(
          (entry) => entry.subtaskId === subtask.id,
        );
        // TASKS.md #195 — look this subtask up in the FULL failure list, not
        // just the single most-relevant one. A wave of many independent
        // parallel subtasks (e.g. 12 month-sheet creates with no dependsOn
        // between them) can have several genuinely fail at once; matching
        // only `failedSubtask` left every failure but one with no recorded
        // reason at all in this run's persisted state.
        const failed = waveResult.failedSubtasks.find((entry) => entry.subtaskId === subtask.id);
        return {
          subtaskId: subtask.id,
          actions: completed?.actions ?? [],
          completed: Boolean(completed),
          verified: completed?.verified,
          failedReason: failed?.reason,
        };
      }),
    );

    // A wave that produced nothing must not emit an empty Accept card — it is
    // a failure to report, not a step to approve. SD-4 says continue rather
    // than abort, so the run advances with this wave marked skipped.
    if (waveResult.actions.length === 0) {
      const reason =
        waveResult.failedSubtask?.reason ?? 'This step produced no changes to apply';
      this.logger.warn(`Stepwise run ${run.runId} wave ${next.waveIndex} empty: ${reason}`);
      emit('status', { message: `Step skipped — ${reason}` });
      await this.agentRunState.applyDecision(run, 'skipped');
      await this.executeStepwiseWave(run, reply, emit, sseEmitter, telemetry);
      return;
    }

    // TASKS.md #195 — a wave that DID produce some actions can still have lost
    // OTHER independent subtasks silently (e.g. 3 of 12 month sheets built,
    // 9 failed) — the Accept card only ever described what succeeded. Surface
    // the gap the same way the one-shot path's `undeliveredSubtasks` already
    // does, so "only got one sheet" has a visible, honest explanation instead
    // of looking like the request was simply under-specified.
    if (waveResult.failedSubtasks.length > 0) {
      const missing = waveResult.failedSubtasks;
      this.logger.warn(
        `Stepwise run ${run.runId} wave ${next.waveIndex}: ${missing.length} of ${next.subtasks.length} ` +
          `subtask(s) failed — ${missing.map((m) => m.subtaskId).join(', ')}`,
      );
      emit('status', {
        message:
          missing.length === 1
            ? `Note: 1 of ${next.subtasks.length} planned steps in this batch produced no changes — ${missing[0].reason.slice(0, 110)}`
            : `Note: ${missing.length} of ${next.subtasks.length} planned steps in this batch produced no changes (e.g. ${missing[0].reason.slice(0, 90)})`,
      });
    }

    const changeSet = await this.changeSetService.createPreview({
      conversationId: run.conversationId,
      traceId: run.traceId,
      prompt: run.prompt,
      context: run.context as AgentWorkbookContext,
      actions: waveResult.actions,
      provenance: {
        sourceRefs: buildWorkbookSourceRefsFromActions(
          waveResult.actions,
          run.context.activeSheetName || 'workbook',
          run.context.activeSheetName,
        ),
        workbookId: run.context.activeSheetName || 'workbook',
        activeSheetName: run.context.activeSheetName,
      },
    });

    const label = this.describeProgressiveWave(waveResult.actions);
    const previousChangeSetId = run.changeSetIds[run.changeSetIds.length - 1];

    emit('actions', {
      actions: waveResult.actions,
      explanation: label,
      userFacingSummary: buildUserFacingSummary({
        answer: label,
        actions: waveResult.actions,
        changes: changeSet.changes,
        activeSheetName: run.context.activeSheetName,
        planSubtasks: next.subtasks.map((subtask) => ({
          id: subtask.id,
          description: subtask.description,
          targetSheet: subtask.targetSheet,
        })),
      }),
      internalDetails: buildInternalDetails({
        tier: 3,
        model: telemetry.model,
        processingLabel: label,
        actions: waveResult.actions,
        legacyExplanation: label,
      }),
      changeSetId: changeSet.changeSetId,
      changes: changeSet.changes,
      irreversibleActionTypes: changeSet.irreversibleActionTypes,
      tier: 3,
      stepIndex: next.waveIndex + 1,
      stepTotal: run.waveTotal,
      stepLabel: label,
      stepwise: true,
      runId: run.runId,
      ...(previousChangeSetId ? { dependsOnChangeSetId: previousChangeSetId } : {}),
    });

    run.changeSetIds.push(changeSet.changeSetId);
    await run.save();
    this.logWorkflowChangeSet(
      run.traceId,
      changeSet.changeSetId,
      waveResult.actions,
      changeSet.changes.length,
    );

    // The stream ends here but the RUN does not — this is what tells the client
    // to accept and then call /continue, rather than treating the build as done.
    emit('wave_ready', {
      runId: run.runId,
      waveIndex: next.waveIndex,
      waveTotal: run.waveTotal,
      hasMore: true,
      changeSetId: changeSet.changeSetId,
    });
    endSseResponse(reply);
  }

  /** Closes out a run whose waves are all decided, reporting skips honestly. */
  private async finishStepwiseRun(
    run: AgentRunDocument,
    reply: FastifyReply,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<void> {
    const skipped = this.agentRunState.summarizeSkipped(run);
    await this.agentRunState.markStatus(run, 'completed');
    await this.markCompleted(run.conversationId);

    // An incomplete build reported as complete is the false-completeness
    // failure CODEBASE_ANALYSIS.md §3.7 keeps re-teaching — say what was left.
    const summary =
      skipped.length === 0
        ? 'All steps applied.'
        : skipped.length === 1
          ? `Done — 1 step was not applied: ${skipped[0].description}`
          : `Done — ${skipped.length} steps were not applied (e.g. ${skipped[0].description})`;

    emit('conversation_end', {
      summary,
      tier: 3,
      runId: run.runId,
      skippedSubtasks: skipped,
    });
    emit('wave_ready', {
      runId: run.runId,
      waveIndex: run.waveIndex,
      waveTotal: run.waveTotal,
      hasMore: false,
    });
    this.logger.log(
      `Stepwise run ${run.runId} complete waves=${run.waveTotal} skipped=${skipped.length}`,
    );
    endSseResponse(reply);
  }

  /**
   * `POST /excel-ai/conversation/continue` — records the client's decision on
   * the wave just emitted and generates the next one (STEPWISE_EXECUTION.md §3).
   *
   * This is the only place wave N+1's Executor can be reached, which is what
   * enforces SD-3's no-look-ahead rule structurally rather than by convention.
   */
  async continueRun(
    body: ContinueRunDto,
    reply: FastifyReply,
    traceId: string | undefined,
    userId?: string,
  ): Promise<void> {
    const run = await this.agentRunState.loadRunForUser(body.runId, userId);

    initSseResponse(reply);
    const emit = (event: string, data: Record<string, unknown>) =>
      writeSseEvent(reply, event, { ...data, conversationId: run.conversationId });
    const sseEmitter = new SseEmitter(emit);
    const telemetry: LlmCallTelemetry = { provider: 'openrouter', modelTier: 'high' };

    try {
      if (run.status === 'completed' || run.status === 'abandoned') {
        emit('conversation_end', { summary: 'This build has already finished.', tier: 3 });
        emit('wave_ready', {
          runId: run.runId,
          waveIndex: run.waveIndex,
          waveTotal: run.waveTotal,
          hasMore: false,
        });
        endSseResponse(reply);
        return;
      }

      // Readback first: the next wave must plan against the OBSERVED result of
      // the wave just accepted, not the shadow workbook's prediction of it.
      if (body.decision === 'accepted') {
        await this.agentRunState.applyReadback(
          run,
          body.readback as AgentWorkbookContext['sheets'] | undefined,
        );
      }

      const { cascadeSkipped } = await this.agentRunState.applyDecision(
        run,
        body.decision as WaveDecision,
      );
      if (cascadeSkipped.length > 0) {
        emit('status', {
          message: `Skipping ${cascadeSkipped.length} step(s) that depended on the step you did not accept.`,
        });
      }

      await this.executeStepwiseWave(run, reply, emit, sseEmitter, telemetry);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Continue failed';
      this.logger.warn(`Stepwise continue failed run=${run.runId} error="${message}"`);
      await this.agentRunState.markStatus(run, 'failed').catch(() => undefined);
      emit('error', { message });
      endSseResponse(reply);
    }
  }

  private async streamWithOrchestrator(
    request: ConversationRequestDto,
    reply: FastifyReply,
    conversationId: string,
    traceId: string,
    history: ConversationMessageEntry[],
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    emit: (event: string, data: Record<string, unknown>) => void,
    routerAssumption?: string,
    complexity?: 0 | 1 | 2 | 3,
    /** Reports the finished action count back for routing telemetry. TASKS.md #165. */
    reportActionCount?: (count: number) => void,
    /** Owner of any stepwise run this request starts — never from the body. */
    userId?: string,
  ): Promise<void> {
    const startedAt = Date.now();
    const intent = classifyIntent(request.message);
    const telemetry: LlmCallTelemetry = { provider: 'openrouter', modelTier: 'high' };
    let success = false;
    let actionsCount: number | undefined;

    const richWorkbookContext = resolveWorkbookContext(request, analysis, request.sheetData);
    const agentContext = buildAgentWorkbookContext(
      richWorkbookContext,
      request.sheetData,
      analysis,
    );
    const basePrompt =
      request.promptContext ?? richWorkbookContext.prompt_context ?? undefined;
    const { enrichedContext, promptContext } = this.enrichAgentContext(
      agentContext,
      basePrompt,
      history,
      request.message,
    );

    const conversationHistory = resolveConversationHistory(request, history).map((entry) => ({
      role: entry.role as 'user' | 'assistant',
      content: entry.content,
    }));

    const sseEmitter = new SseEmitter(emit);

    try {
      // TASKS.md #174 — progressive emission ("B"). Turn each finished
      // execution wave into an Accept card immediately instead of holding every
      // card until the whole run returns.
      //
      // Only creation and plain-content actions go out early: the consolidation,
      // chart and presentation passes rewrite formulas, chart anchors and
      // formatting once they can see the whole build, and a card must never
      // promise an action a later pass will change (see progressive-emit.util).
      //
      // Everything shown early is recorded by structural key so the final
      // emission — which still runs finalizeActions over ALL actions, because
      // the global passes need that — can exclude it instead of double-showing.
      const progressive = {
        emittedKeys: [] as string[],
        cardCount: 0,
        lastChangeSetId: undefined as string | undefined,
        onWaveComplete: undefined as
          | ((waveActions: SheetAction[], waveIndex: number) => Promise<void>)
          | undefined,
      };

      progressive.onWaveComplete = async (waveActions) => {
        const early = selectEarlyEmittable(waveActions);
        if (early.length === 0) return;

        // One card per phase, never a mixed one — TASKS.md #175.
        for (const group of splitEarlyByPhase(early)) {
          await emitProgressiveCard(group);
        }
      };

      const emitProgressiveCard = async (early: SheetAction[]) => {
        const changeSet = await this.changeSetService.createPreview({
          conversationId,
          traceId,
          prompt: request.message,
          context: enrichedContext,
          actions: early,
          provenance: {
            sourceRefs: buildWorkbookSourceRefsFromActions(
              early,
              enrichedContext.activeSheetName || 'workbook',
              enrichedContext.activeSheetName,
            ),
            workbookId: enrichedContext.activeSheetName || 'workbook',
            activeSheetName: enrichedContext.activeSheetName,
          },
        });

        progressive.cardCount += 1;
        const label = this.describeProgressiveWave(early);

        emit('actions', {
          actions: early,
          explanation: label,
          userFacingSummary: buildUserFacingSummary({
            answer: label,
            actions: early,
            changes: changeSet.changes,
            activeSheetName: enrichedContext.activeSheetName,
          }),
          internalDetails: buildInternalDetails({
            tier: 3,
            model: telemetry.model,
            processingLabel: label,
            actions: early,
            legacyExplanation: label,
          }),
          changeSetId: changeSet.changeSetId,
          changes: changeSet.changes,
          irreversibleActionTypes: changeSet.irreversibleActionTypes,
          tier: 3,
          stepLabel: label,
          // Deliberately no stepIndex/stepTotal: the total is unknown while the
          // run is still going, and a "Step 1 of ?" badge would be a worse lie
          // than none. The final emission carries the real numbering.
          progressive: true,
          ...(progressive.lastChangeSetId
            ? { dependsOnChangeSetId: progressive.lastChangeSetId }
            : {}),
        });

        progressive.lastChangeSetId = changeSet.changeSetId;
        progressive.emittedKeys.push(...keysFor(early));
        this.logWorkflowChangeSet(traceId, changeSet.changeSetId, early, changeSet.changes.length);
      };

      // TASKS.md #153 / STEPWISE_EXECUTION.md — plan the whole build, then
      // execute and preview only the FIRST wave, ending the stream so the
      // client can accept before anything downstream is generated (SD-3).
      // Falls through to the one-shot path below whenever the flag is off or
      // the plan turns out to be a single wave (nothing to gate).
      if (isStepwiseExecutionEnabled()) {
        const handled = await this.tryStartStepwiseRun({
          request,
          reply,
          conversationId,
          traceId,
          emit,
          sseEmitter,
          enrichedContext,
          promptContext,
          conversationHistory,
          routerAssumption,
          complexity,
          telemetry,
          userId,
          startedAt,
        });
        if (handled) {
          success = true;
          return;
        }
      }

      const orchestratorResult = await this.orchestrator.runDetailed(
        {
          prompt: request.message,
          context: enrichedContext,
          conversationHistory,
          promptContext,
          conversationId,
          correlationId: traceId,
          toolEmit: emit,
          routerAssumption,
          complexity: complexity ?? 3,
          onWaveComplete: progressive.onWaveComplete,
        },
        sseEmitter,
        // Populates telemetry.usage/model from the real Planner+Executor+Verifier
        // calls this run makes — previously never wired, so every Tier-3
        // audit_logs row here always reported promptTokens/completionTokens: 0.
        telemetry,
      );
      const rawActions = orchestratorResult.actions;

      if (orchestratorResult.clarificationRequested) {
        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: '[Clarification needed]',
          type: 'clarification',
          timestamp: new Date(),
        });
        emit('done', { message: 'awaiting_clarification' });
        endSseResponse(reply);
        success = true;
        return;
      }

      if (!orchestratorResult.verifierPassed) {
        if (orchestratorResult.partialProgress && rawActions.length > 0) {
          // TASKS.md #195 — describe EVERY failed subtask, not just the first.
          // A wave of independent parallel subtasks (e.g. many month-sheet
          // creates with no dependsOn between them) can have several
          // genuinely fail at once; reporting only `failedSubtask` understated
          // how much of the request actually failed.
          const failedReason =
            orchestratorResult.failedSubtasks.length > 1
              ? `${orchestratorResult.failedSubtasks.length} steps could not be completed, including: ${orchestratorResult.failedSubtasks[0].reason}`
              : orchestratorResult.failedSubtask?.reason ??
                'A later step could not be completed';
          const finalized = this.engine.finalizeActions(
            rawActions,
            analysis,
            richWorkbookContext,
            request.message,
            enrichedContext.priorTurnActions,
            // TASKS.md #172 — this argument was missing here while the success
            // path passed it, so a partial run silently lost the client's
            // probed capabilities and fell back to the non-dynamic-array
            // consolidation even on a host that supports it (#152).
            request.excelCapabilities,
          );

          // PHASE-ORDER the actions before they are previewed — TASKS.md #172.
          //
          // The success path gets this from `createActionWaveChangeSets` ->
          // `splitIntoActionWaves`, which buckets create -> content -> formula
          // -> format -> layout -> chart. This branch built a preview straight
          // from `finalizeActions`, which does NOT reorder, so actions reached
          // Excel in the order the plan happened to accumulate them.
          //
          // That ordering is actively hostile here: planner.prompt.ts rule 0
          // requires the Main-sheet subtasks to be emitted FIRST and the twelve
          // month subtasks LAST (a token-budget rule, so a truncated plan loses
          // the boilerplate rather than the dashboard). Correct for planning,
          // wrong for applying — it puts `=SUM(July!H:H)` on Main ahead of the
          // ADD_SHEET that creates July. The dependency graph orders EXECUTION
          // against the shadow workbook; nothing was ordering the real write.
          //
          // Flattening the waves keeps this branch's single-card UX while
          // giving it the same ordering guarantee the success path has.
          const actions = splitIntoActionWaves(finalized).flatMap((wave) => wave.actions);
          actionsCount = actions.length;
          reportActionCount?.(actions.length);

          const answer =
            `I completed **${orchestratorResult.completedSubtasks.length}** step(s) and prepared **${actions.length}** change(s) for preview, ` +
            `but could not finish the full request: ${failedReason}. ` +
            `Want me to retry just that step?`;

          const changeSet = await this.changeSetService.createPreview({
            conversationId,
            traceId,
            prompt: request.message,
            context: enrichedContext,
            actions,
            provenance: {
              sourceRefs: buildWorkbookSourceRefsFromActions(
                actions,
                enrichedContext.activeSheetName || 'workbook',
                enrichedContext.activeSheetName,
              ),
              workbookId: enrichedContext.activeSheetName || 'workbook',
              activeSheetName: enrichedContext.activeSheetName,
            },
          });

          const processingLabel =
            'Partial progress: earlier steps succeeded; a later step failed. Review and accept what is ready, or retry the failed step.';
          const userFacingSummary = buildUserFacingSummary({
            answer,
            actions,
            changes: changeSet.changes,
            assumption: routerAssumption,
            activeSheetName: enrichedContext.activeSheetName,
          });
          const internalDetails = buildInternalDetails({
            tier: 3,
            model: telemetry.model,
            processingLabel,
            assumption: routerAssumption,
            actions,
            legacyExplanation: processingLabel,
          });

          await this.saveMessage(conversationId, {
            id: `msg_${Date.now()}_assistant`,
            role: 'assistant',
            content: answer,
            type: 'answer',
            timestamp: new Date(),
            metadata: {
              actions,
              changeSetId: changeSet.changeSetId,
              partialProgress: true,
              failedSubtask: orchestratorResult.failedSubtask,
              failedSubtasks: orchestratorResult.failedSubtasks,
            },
          });

          emit('answer', { answer });
          emit('actions', {
            actions,
            explanation: processingLabel,
            userFacingSummary,
            internalDetails,
            changeSetId: changeSet.changeSetId,
            changes: changeSet.changes,
            irreversibleActionTypes: changeSet.irreversibleActionTypes,
            partialProgress: true,
            failedSubtask: orchestratorResult.failedSubtask,
            failedSubtasks: orchestratorResult.failedSubtasks,
            tier: 3,
          });
          emit('conversation_end', {
            summary: 'Partial changes ready — review and accept, or retry the failed step.',
          });
          await this.markCompleted(conversationId);
          this.logWorkflowChangeSet(
            traceId,
            changeSet.changeSetId,
            actions,
            changeSet.changes.length,
          );
          this.finalizeWorkflow(traceId, 'awaiting_accept', {
            changeSetId: changeSet.changeSetId,
            route: 'write',
            tier: 3,
            durationMs: Date.now() - startedAt,
            sseOutput: {
              partialProgress: true,
              changeSetId: changeSet.changeSetId,
              failedSubtask: orchestratorResult.failedSubtask,
              actionTypes: actions.map((a) => a.type),
            },
          });
          endSseResponse(reply);
          success = true;
          return;
        }

        const answer =
          'I could not complete and verify the full request, so no partial changes were sent to Excel. Please retry or split the request into smaller steps.';
        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: answer,
          type: 'answer',
          timestamp: new Date(),
        });
        emit('answer', { answer });
        emit('conversation_end', { summary: 'No unverified changes were applied.' });
        await this.markCompleted(conversationId);
        this.finalizeWorkflow(traceId, 'failed', {
          route: 'write',
          tier: 3,
          durationMs: Date.now() - startedAt,
          sseOutput: { answer, verifierPassed: false },
        });
        endSseResponse(reply);
        success = true;
        return;
      }

      if (rawActions.length === 0) {
        this.assertWriteRouteProducedActions({
          conversationId,
          message: request.message,
          actionsLength: 0,
        });
      }

      const finalizedActions = this.engine.finalizeActions(
        rawActions,
        analysis,
        richWorkbookContext,
        request.message,
        enrichedContext.priorTurnActions,
        request.excelCapabilities,
      );

      // TASKS.md #174 — finalize still runs over EVERY action, because the
      // consolidation, chart and presentation passes need the whole build in
      // view. Anything already shown as a progressive card is removed here so
      // it is not offered twice; what remains are the phases those passes own
      // plus anything no early card covered.
      const actions = excludeAlreadyEmitted(finalizedActions, progressive.emittedKeys);

      if (progressive.cardCount > 0) {
        this.logger.log(
          `Progressive emission: ${progressive.cardCount} card(s) sent during the run ` +
            `(${progressive.emittedKeys.length} actions); ${actions.length} of ` +
            `${finalizedActions.length} finalized actions remain for the closing cards.`,
        );
      }

      actionsCount = finalizedActions.length;
      reportActionCount?.(finalizedActions.length);

      // Asserted against the FULL finalized list: a run whose entire output was
      // already delivered progressively is a success, not an empty write.
      this.assertWriteRouteProducedActions({
        conversationId,
        message: request.message,
        actionsLength: finalizedActions.length,
      });

      const answer = `I'll apply the prepared changes to your sheet.`;

      // Large multi-sheet builds ("a sheet per month, then fill each in") split
      // into staged accept waves — sheet creates reviewed/accepted before the
      // writes that depend on them exist as their own card. A pure-write or
      // pure-create batch (the common case) comes back as a single wave,
      // identical to today's behavior.
      // Everything was already delivered progressively — emitting the closing
      // set would render a "0 changes ready for review" card. TASKS.md #174.
      if (actions.length === 0 && progressive.cardCount > 0) {
        this.logger.log(
          `Progressive emission delivered the entire build in ${progressive.cardCount} card(s); no closing card needed.`,
        );
        emit('conversation_end', { summary: 'Review changes and accept or reject.', tier: 3 });
        await this.markCompleted(conversationId);
        this.finalizeWorkflow(traceId, 'awaiting_accept', {
          route: 'write',
          tier: 3,
          durationMs: Date.now() - startedAt,
        });
        endSseResponse(reply);
        success = true;
        return;
      }

      const waveChangeSets = await this.createActionWaveChangeSets(actions, {
        conversationId,
        traceId,
        prompt: request.message,
        context: enrichedContext,
      });
      const lastChangeSet = waveChangeSets[waveChangeSets.length - 1].changeSet;
      const combinedChangesLength = waveChangeSets.reduce(
        (sum, w) => sum + w.changeSet.changes.length,
        0,
      );

      const processingLabel = tierProcessingLabel(3);

      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: answer,
        type: 'answer',
        timestamp: new Date(),
        metadata: this.buildWriteMetadata(actions, lastChangeSet.changeSetId),
      });

      emit('answer', { answer, tier: 3 });

      // TASKS.md #155 — a plan whose Executor delivered nothing for some of its
      // own subtasks is incomplete, and saying so is the §3.7 rule this
      // codebase keeps re-learning. The Accept card already excludes these from
      // its promises; this makes the omission visible rather than merely quiet.
      // TASKS.md #171 — the build proceeded under an assumption; say so. These
      // are questions the Planner raised and we deliberately did NOT block on,
      // so hiding them would be the §3.7 false-completeness shape: the user
      // would see a finished workbook and never learn a guess was made.
      if (orchestratorResult.openQuestions.length > 0) {
        const asked = orchestratorResult.openQuestions;
        this.logger.log(
          `Proceeded under ${asked.length} open assumption(s) rather than blocking: ${asked.join(' | ')}`,
        );
        emit('status', {
          message:
            `I built this using my best reading of your request. ${asked.length === 1 ? 'One thing' : `${asked.length} things`} to confirm: ` +
            asked.map((q) => q.trim()).join(' '),
        });
      }

      if (orchestratorResult.undeliveredSubtasks.length > 0) {
        const missing = orchestratorResult.undeliveredSubtasks;
        this.logger.warn(
          `Plan/delivery gap: ${missing.length} planned subtask(s) produced no actions — ` +
            missing.map((m) => `${m.id} (${m.targetSheet})`).join(', '),
        );
        emit('status', {
          message:
            missing.length === 1
              ? `Note: 1 planned step produced no changes — ${missing[0].description.slice(0, 110)}`
              : `Note: ${missing.length} planned steps produced no changes (e.g. ${missing[0].description.slice(0, 90)})`,
        });
      }

      // TASKS.md #167 — map each finalized action back to the subtask that
      // produced it, so every staged step can describe its OWN work. The
      // provenance was always in `completedSubtasks`; `finalizeActions` takes a
      // flat array and dropped it (CODEBASE_ANALYSIS.md §3.7's shape again).
      const actionAttribution = attributeActionsToSubtasks(
        actions,
        orchestratorResult.completedSubtasks ?? [],
      );

      // Seeded from the last progressive card so the closing cards cannot be
      // accepted before the content they depend on. Without this the chain
      // restarts and a formatting card could apply to sheets that do not exist
      // yet — the #80 dependency guard would refuse it, but as a confusing
      // failure rather than a disabled button. TASKS.md #174.
      let previousChangeSetId: string | undefined = progressive.lastChangeSetId;
      let firstUserFacingSummary: ReturnType<typeof buildUserFacingSummary> | undefined;
      for (const [waveIndex, { wave, changeSet }] of waveChangeSets.entries()) {
        const isFirstWave = !previousChangeSetId;
        const userFacingSummary = buildUserFacingSummary({
          answer: isFirstWave ? answer : wave.label,
          actions: wave.actions,
          changes: changeSet.changes,
          assumption: isFirstWave ? routerAssumption : undefined,
          activeSheetName: enrichedContext.activeSheetName,
          // Plan intent, scoped to THIS step's own actions.
          //
          // TASKS.md #161 had to disable intent entirely for staged builds
          // because every card rendered the same whole-plan bullets — "Create
          // 13 sheets" promising the formulas and charts three steps away,
          // which is #155's over-promising in a new shape. That fix was
          // correct and explicitly temporary: it left every staged card
          // describing machinery ("83 formatting changes") when the data
          // needed to describe work existed one layer up.
          //
          // Now each wave reports only the subtasks its own actions came from,
          // so the over-promising is impossible by construction rather than by
          // suppression. A wave that attributes to nothing (pure pass-generated
          // formatting on unowned sheets) yields [] and correctly falls back to
          // #140's action rollup — never an empty card, the bug #149 already
          // hit once. TASKS.md #167.
          planSubtasks:
            waveChangeSets.length === 1
              ? orchestratorResult.planSubtasks
              : (() => {
                  const scoped = intentForWave(
                    wave.actionIndexes,
                    actionAttribution,
                    orchestratorResult.planSubtasks,
                  );
                  return scoped.length > 0 ? scoped : undefined;
                })(),
        });
        firstUserFacingSummary ??= userFacingSummary;
        const internalDetails = buildInternalDetails({
          tier: 3,
          model: telemetry.model,
          processingLabel: isFirstWave ? processingLabel : wave.label,
          assumption: isFirstWave ? routerAssumption : undefined,
          actions: wave.actions,
          legacyExplanation: isFirstWave ? processingLabel : wave.label,
        });

        emit('actions', {
          actions: wave.actions,
          explanation: isFirstWave ? processingLabel : wave.label,
          userFacingSummary,
          internalDetails,
          changeSetId: changeSet.changeSetId,
          changes: changeSet.changes,
          irreversibleActionTypes: changeSet.irreversibleActionTypes,
          tier: 3,
          // Position in a staged build. The thing TASKS.md #141's two-wave
          // split lacked: without it, accepting step 1 and stopping left a
          // half-built workbook that looked finished. TASKS.md #160.
          stepIndex: waveIndex + 1,
          stepTotal: waveChangeSets.length,
          stepLabel: wave.label,
          // Gate: the frontend must not let this wave's Accept fire until the
          // wave named here has been accepted (its sheets/ranges must exist).
          ...(previousChangeSetId ? { dependsOnChangeSetId: previousChangeSetId } : {}),
        });

        this.logWorkflowChangeSet(traceId, changeSet.changeSetId, wave.actions, changeSet.changes.length);
        previousChangeSetId = changeSet.changeSetId;
      }

      emit('conversation_end', { summary: 'Review changes and accept or reject.', tier: 3 });
      await this.markCompleted(conversationId);
      this.finalizeWorkflow(traceId, 'awaiting_accept', {
        changeSetId: lastChangeSet.changeSetId,
        route: 'write',
        tier: 3,
        durationMs: Date.now() - startedAt,
        sseOutput: {
          answer,
          changeSetId: lastChangeSet.changeSetId,
          actionTypes: actions.map((a) => a.type),
          changesLength: combinedChangesLength,
          userFacingSummary: firstUserFacingSummary,
        },
      });
      endSseResponse(reply);
      success = true;

      this.logger.log(
        `Orchestrator response trace=${traceId} conversation=${conversationId} actions=${actions.length} durationMs=${Date.now() - startedAt}`,
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Orchestrator failed';
      this.logger.warn(
        `Orchestrator failed trace=${traceId} conversation=${conversationId} durationMs=${Date.now() - startedAt} error="${this.clipForLog(message, 300)}"`,
      );
      // End the SSE stream cleanly — do not rethrow. Rethrowing after parallel
      // LLM aborts can surface as unhandled TypeError("terminated") and crash nodemon.
      emit('error', {
        message:
          error instanceof LlmRequestError
            ? `AI provider failed (${error.status}): ${message}`
            : message,
      });
      endSseResponse(reply);
    } finally {
      await this.auditService.logLLMCall({
        traceId,
        model: telemetry.model ?? 'orchestrator',
        tier: (telemetry.modelTier ?? 'high') as LLMTier,
        intent,
        promptTokens: telemetry.usage?.promptTokens ?? 0,
        completionTokens: telemetry.usage?.completionTokens ?? 0,
        latencyMs: Date.now() - startedAt,
        success,
        actionsCount,
      });
    }
  }

  private async streamWithPlanner(
    request: ConversationRequestDto,
    reply: FastifyReply,
    conversationId: string,
    traceId: string,
    history: ConversationMessageEntry[],
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    emit: (event: string, data: Record<string, unknown>) => void,
    complexity: 0 | 1 | 2 | 3 = 3,
  ): Promise<void> {
    const startedAt = Date.now();
    const intent = classifyIntent(request.message);
    const telemetry: LlmCallTelemetry = { provider: 'openrouter', modelTier: 'high' };
    let success = false;

    const richWorkbookContext = resolveWorkbookContext(request, analysis, request.sheetData);
    const agentContext = buildAgentWorkbookContext(richWorkbookContext, request.sheetData, analysis);
    const basePrompt = request.promptContext ?? richWorkbookContext.prompt_context ?? undefined;
    const { enrichedContext, promptContext } = this.enrichAgentContext(
      agentContext,
      basePrompt,
      history,
      request.message,
    );
    const conversationHistory = resolveConversationHistory(request, history).map((entry) => ({
      role: entry.role as 'user' | 'assistant',
      content: entry.content,
    }));

    emit('thinking', { message: '🧠 Building a step-by-step plan across your workbook…' });

    try {
      const plan = await this.orchestrator.planOnly(
        {
          prompt: request.message,
          context: enrichedContext,
          conversationHistory,
          promptContext,
          conversationId,
          correlationId: traceId,
          complexity,
        },
        // Same wiring as streamWithOrchestrator — this path only calls the
        // Planner, but previously reported promptTokens/completionTokens: 0
        // for the same reason (telemetry.usage was never populated).
        telemetry,
      );

      if (plan.clarificationsNeeded.length > 0) {
        const question = plan.clarificationsNeeded.join(' ');
        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: `[Clarification needed]: ${question}`,
          type: 'clarification',
          timestamp: new Date(),
        });
        emit('clarification', { question, suggestions: [], ambiguityScore: 0 });
        emit('done', { message: 'awaiting_clarification' });
        this.finalizeWorkflow(traceId, 'clarifying', {
          route: 'write',
          sseOutput: { clarification: question },
        });
        endSseResponse(reply);
        success = true;
        return;
      }

      const planPayload = this.buildPlanPayload(plan, request.message, enrichedContext);
      const answer =
        planPayload.steps.length > 0
          ? `Here's a ${planPayload.steps.length}-step plan. Review it, then run it as an action to preview and apply the changes.`
          : 'I could not break this request into concrete steps. Try rephrasing, or switch to Action mode.';

      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: answer,
        type: 'answer',
        timestamp: new Date(),
      });

      emit('answer', { answer });
      emit('plan', planPayload as unknown as Record<string, unknown>);
      emit('conversation_end', { summary: 'Plan ready — run as action to apply.' });
      await this.markCompleted(conversationId);
      this.finalizeWorkflow(traceId, 'completed', {
        route: 'write',
        durationMs: Date.now() - startedAt,
        sseOutput: {
          kind: 'plan',
          answer,
          stepCount: planPayload.steps.length,
        },
      });
      endSseResponse(reply);
      success = true;

      this.logger.log(
        `Planner response trace=${traceId} conversation=${conversationId} steps=${planPayload.steps.length} durationMs=${Date.now() - startedAt}`,
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Planner failed';
      this.logger.warn(
        `Planner failed trace=${traceId} conversation=${conversationId} durationMs=${Date.now() - startedAt} error="${this.clipForLog(message, 300)}"`,
      );
      emit('error', { message });
      endSseResponse(reply);
    } finally {
      await this.auditService.logLLMCall({
        traceId,
        model: telemetry.model ?? 'planner',
        tier: (telemetry.modelTier ?? 'high') as LLMTier,
        intent,
        promptTokens: telemetry.usage?.promptTokens ?? 0,
        completionTokens: telemetry.usage?.completionTokens ?? 0,
        latencyMs: Date.now() - startedAt,
        success,
      });
    }
  }

  private buildPlanPayload(
    plan: PlannerOutput,
    prompt: string,
    context: AgentWorkbookContext,
  ): {
    prompt: string;
    summary: string;
    steps: { title: string; detail?: string }[];
    affectedSheets: string[];
    estimatedRows: number;
    safestApproach: string;
  } {
    const steps = plan.subtasks.map((subtask) => ({
      title: subtask.description,
      detail: subtask.targetSheet ? `Sheet: ${subtask.targetSheet}` : undefined,
    }));

    const affectedSheets = Array.from(
      new Set(plan.subtasks.map((s) => s.targetSheet).filter((name): name is string => Boolean(name))),
    );

    const estimatedRows = affectedSheets.reduce((sum, name) => {
      const sheet = context.sheets.find((s) => s.name === name);
      const dataRows = sheet ? Math.max(0, sheet.rowCount - 1) : 0;
      return sum + dataRows;
    }, 0);

    const safestApproach =
      plan.reasoning?.trim() ||
      'Review the plan, then run it as an action to preview every change before applying.';

    return {
      prompt,
      summary: `${plan.subtasks.length} step(s) · confidence ${plan.confidence}`,
      steps,
      affectedSheets,
      estimatedRows,
      safestApproach,
    };
  }

  private async streamWithOpenAi(
    request: ConversationRequestDto,
    reply: FastifyReply,
    conversationId: string,
    traceId: string,
    history: ConversationMessageEntry[],
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<void> {
    let fullText = '';
    const telemetry: LlmCallTelemetry = {};
    const startedAt = Date.now();
    let success = false;
    let errorCode: string | undefined;
    let actionsCount: number | undefined;
    const intent = classifyIntent(request.message);
    const richWorkbookContext = resolveWorkbookContext(request, analysis, request.sheetData);
    const conversationTurns = resolveConversationHistory(request, history);
    const llmPlan = this.engine.planLlmCall(
      request.message,
      request.sheetData,
      analysis,
      history,
      resolveEngineWorkbookMeta(request),
      richWorkbookContext,
    );

    const readOnly = modeIsReadOnly(request.mode);
    if (readOnly && llmPlan.messages[0]?.role === 'system') {
      const directive =
        request.mode === 'plan' ? PLAN_MODE_DIRECTIVE : ASK_MODE_READONLY_DIRECTIVE;
      llmPlan.messages[0] = {
        ...llmPlan.messages[0],
        content: `${llmPlan.messages[0].content}\n\n${directive}`,
      };
    }

    emit('thinking', { message: `🧠 ${llmPlan.thinkingMessage}` });

    try {
      for await (const token of this.engine.streamPlannedLlm(llmPlan, telemetry)) {
        fullText += token;
      }
      success = true;

      const structured = this.engine.parseStructuredResponse(
        fullText,
        analysis,
        request.message,
        richWorkbookContext,
      );
      // Spec 23: strip internal vocabulary / mode-switch pitches from ask/plan LLM copy.
      if (readOnly && structured?.kind === 'answer') {
        structured.answer = sanitizeAskAnswer(structured.answer);
      }
      if (readOnly && structured?.kind === 'actions') {
        structured.answer = sanitizeAskAnswer(structured.answer);
      }
      // Task #90: an answer that contradicts its own arithmetic (pre-tax + tax !=
      // total) must not be presented as fact. Applies in every mode — a wrong
      // figure is just as damaging when it accompanies a write.
      if (structured?.kind === 'answer' || structured?.kind === 'actions') {
        const consistency = annotateAnswerConsistency(structured.answer);
        if (consistency.issue) {
          this.logger.warn(
            `Answer self-inconsistency trace=${traceId} conversation=${conversationId} ` +
              `kind=${consistency.issue.kind} parts=${consistency.issue.parts.join('/')} ` +
              `expected=${consistency.issue.expected} stated=${consistency.issue.stated}`,
          );
          structured.answer = consistency.answer;
        }
      }
      const fallbackText = fullText.trim() || 'I could not generate a response.';
      if (structured?.kind === 'actions') {
        actionsCount = structured.actions.length;
      }
      this.logger.log(
        `AI response trace=${traceId} conversation=${conversationId} called=true provider=${telemetry.provider ?? 'unknown'} modelTier=${telemetry.modelTier ?? 'unknown'} model=${telemetry.model ?? 'unknown'} tokens=${this.formatUsage(telemetry)} durationMs=${Date.now() - startedAt} response="${this.clipForLog(fallbackText)}"`,
      );
      // Task #92: the line above is the ONLY record of provider/model/token counts,
      // and logger.log output reaches no file logger — task #86 could not be
      // root-caused until the user pasted it from their terminal by hand. Persist it.
      this.recordLlmCallTrace(traceId, telemetry, {
        durationMs: Date.now() - startedAt,
        emptyResponse: !fullText.trim(),
        structuredKind: structured?.kind ?? 'none',
      });

      if (!structured) {
        const deterministicTable = readOnly ? null : tryDeterministicTableCreate(request.message);
        if (deterministicTable) {
          const { plan, actions: tableFallback } = deterministicTable;
          const answer = `Created **${plan.rowCount}** rows with columns: ${plan.headers.join(', ')}.`;
          this.logger.log(
            `Table create (LLM parse fallback) trace=${traceId} conversation=${conversationId}`,
          );
          await this.saveMessage(conversationId, {
            id: `msg_${Date.now()}_assistant`,
            role: 'assistant',
            content: answer,
            type: 'answer',
            timestamp: new Date(),
            metadata: { actions: tableFallback },
          });
          emit('answer', { answer });
          emit('actions', {
            actions: tableFallback,
            explanation: 'Wrote headers and all data rows to your sheet.',
          });
          emit('conversation_end', { summary: 'Changes applied.' });
          await this.markCompleted(conversationId);
          this.finalizeWorkflow(traceId, 'completed', {
            durationMs: Date.now() - startedAt,
            sseOutput: {
              kind: 'actions',
              answer,
              actionTypes: tableFallback.map((a) => a.type),
              source: 'table_fallback',
            },
          });
          endSseResponse(reply);
          return;
        }

        const retryHint =
          'I understood your request but could not parse the AI response. Please try again — e.g. "Generate 10 rows of sample GST purchase data with headers".';
        const rawAnswer =
          fallbackText.length > 20 && !fallbackText.startsWith('{')
            ? `${fallbackText}\n\n${retryHint}`
            : retryHint;
        const answer = readOnly ? sanitizeAskAnswer(rawAnswer) : rawAnswer;
        const parseFailMetadata = await this.buildAnswerPersistMetadata(
          conversationId,
          answer,
          false,
        );

        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: answer,
          type: 'answer',
          timestamp: new Date(),
          ...(parseFailMetadata ? { metadata: parseFailMetadata } : {}),
        });
        emit('answer', { answer });
        emit('conversation_end', { summary: 'Completed.' });
        await this.markCompleted(conversationId);
        this.finalizeWorkflow(traceId, 'completed', {
          durationMs: Date.now() - startedAt,
          sseOutput: {
            kind: 'answer',
            answer,
            parseFailed: true,
            pendingWritePlan: Boolean(parseFailMetadata?.pendingWritePlan),
          },
        });
        endSseResponse(reply);
        return;
      }

    if (structured.kind === 'question') {
      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: structured.question,
        type: 'question',
        timestamp: new Date(),
        metadata: { questionOptions: structured.options },
      });
      emit('question', {
        question: structured.question,
        options: structured.options,
      });
      this.finalizeWorkflow(traceId, 'clarifying', {
        durationMs: Date.now() - startedAt,
        sseOutput: { question: structured.question, options: structured.options },
      });
      endSseResponse(reply);
      return;
    }

    if (structured.kind === 'actions') {
      if (readOnly) {
        // Defense-in-depth: ask/plan modes must never apply edits even if the
        // model produced actions. Strip them and answer only.
        const { removedCount } = stripWriteActions(structured.actions);
        const modeLabel = request.mode === 'plan' ? 'Plan' : 'Ask';
        const note =
          removedCount > 0
            ? `\n\n_${modeLabel} mode is read-only. Want me to apply these changes when you're ready to edit?_`
            : '';
        const answer = sanitizeAskAnswer(`${structured.answer}${note}`);
        const readOnlyMetadata = await this.buildAnswerPersistMetadata(
          conversationId,
          answer,
          false,
        );
        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: answer,
          type: 'answer',
          timestamp: new Date(),
          ...(readOnlyMetadata ? { metadata: readOnlyMetadata } : {}),
        });
        emit('answer', { answer });
        emit('conversation_end', { summary: 'Read-only response.' });
        await this.markCompleted(conversationId);
        this.finalizeWorkflow(traceId, 'completed', {
          route: 'ask',
          durationMs: Date.now() - startedAt,
          sseOutput: { kind: 'answer', answer, readOnly: true },
        });
        endSseResponse(reply);
        return;
      }

      this.assertWriteRouteProducedActions({
        conversationId,
        message: request.message,
        actionsLength: structured.actions.length,
      });

      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: structured.answer,
        type: 'answer',
        timestamp: new Date(),
        metadata: { actions: structured.actions },
      });
      emit('answer', { answer: structured.answer });
      emit('actions', {
        actions: structured.actions,
        explanation: structured.explanation,
      });
      emit('conversation_end', { summary: 'Changes applied.' });
      await this.markCompleted(conversationId);
      this.finalizeWorkflow(traceId, 'completed', {
        durationMs: Date.now() - startedAt,
        sseOutput: {
          kind: 'actions',
          answer: structured.answer,
          actionTypes: structured.actions.map((a) => a.type),
        },
      });
      endSseResponse(reply);
      return;
    }

    // Confirm-only prose with no actions: persist a resumable plan so "yes" can finish the write.
    const pendingMetadata = await this.buildAnswerPersistMetadata(
      conversationId,
      structured.answer,
      false,
    );
    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: structured.answer,
      type: 'answer',
      timestamp: new Date(),
      ...(pendingMetadata ? { metadata: pendingMetadata } : {}),
    });
    emit('answer', { answer: structured.answer });
    emit('conversation_end', { summary: 'Completed.' });
    await this.markCompleted(conversationId);
    this.finalizeWorkflow(traceId, 'completed', {
      durationMs: Date.now() - startedAt,
      sseOutput: {
        kind: 'answer',
        answer: structured.answer,
        pendingWritePlan: Boolean(pendingMetadata?.pendingWritePlan),
      },
    });
    endSseResponse(reply);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'AI provider failed';
      errorCode =
        error instanceof LlmRequestError
          ? String(error.status)
          : error instanceof Error
            ? error.name
            : 'UNKNOWN_ERROR';
      this.logger.warn(
        `AI failed trace=${traceId} conversation=${conversationId} called=true provider=${telemetry.provider ?? 'unknown'} modelTier=${telemetry.modelTier ?? 'unknown'} model=${telemetry.model ?? 'unknown'} durationMs=${Date.now() - startedAt} error="${this.clipForLog(message, 300)}"`,
      );
      throw error;
    } finally {
      await this.auditService.logLLMCall({
        traceId,
        model: telemetry.model ?? 'unknown',
        tier: (telemetry.modelTier ?? 'medium') as LLMTier,
        intent,
        promptTokens: telemetry.usage?.promptTokens ?? 0,
        completionTokens: telemetry.usage?.completionTokens ?? 0,
        latencyMs: Date.now() - startedAt,
        success,
        errorCode,
        actionsCount,
        rawUsage: telemetry.usage
          ? {
              prompt_tokens: telemetry.usage.promptTokens,
              completion_tokens: telemetry.usage.completionTokens,
              total_tokens: telemetry.usage.totalTokens,
            }
          : undefined,
      });
    }
  }

  async handleToolResult(body: {
    conversationId: string;
    requestId: string;
    tool: string;
    values?: unknown[][];
    error?: string;
  }): Promise<{ accepted: boolean }> {
    if (body.tool !== 'get_range_data') {
      return { accepted: false };
    }

    const accepted = this.toolBridge.deliverResult(body.conversationId, body.requestId, {
      values: body.values ?? [],
      error: body.error,
    });

    return { accepted };
  }

  private async applyRefinementContext(
    request: ConversationRequestDto,
  ): Promise<ConversationRequestDto> {
    if (!request.refinementChangeSetId) {
      return request;
    }

    const changeSet = await this.changeSetService.getById(request.refinementChangeSetId);
    if (!changeSet) {
      throw new NotFoundException(`Change set ${request.refinementChangeSetId} not found`);
    }

    const refinement = buildRefinementContext(changeSet);
    const mergedPromptContext = [refinement.promptContext, request.promptContext]
      .filter(Boolean)
      .join('\n\n');

    this.logger.log(
      `Quick edit against change set ${changeSet.changeSetId} (${changeSet.changes.length} cells)`,
    );

    return {
      ...request,
      sheetData: refinement.sheetData,
      workbookContext: refinement.richWorkbookContext,
      promptContext: mergedPromptContext,
      sheetCompression: {
        originalRowCount: refinement.sheetData.length,
        compressedRowCount: refinement.sheetData.length,
        truncated: false,
        onDemandFetchEnabled: true,
      },
    };
  }

  private validateRequest(request: ConversationRequestDto): void {
    if (!request.message?.trim()) {
      throw new BadRequestException('Message is required');
    }

    if (!Array.isArray(request.sheetData)) {
      throw new BadRequestException('Invalid sheet data format');
    }

    const isQuickEdit = Boolean(request.refinementChangeSetId);
    const declaredRowCount = this.resolveDeclaredRowCount(request);
    const effectiveRowCount = Math.max(declaredRowCount, request.sheetData.length);
    const isMetadataFirst =
      isQuickEdit ||
      Boolean(request.sheetCompression?.onDemandFetchEnabled) ||
      (declaredRowCount > request.sheetData.length && request.sheetData.length <= 20);

    if (effectiveRowCount > 10_000) {
      throw new BadRequestException('Sheet too large (max 10000 rows)');
    }

    if (!isMetadataFirst && request.sheetData.length > 1000) {
      throw new BadRequestException('Sheet too large (max 1000 rows)');
    }

    const columnCount = request.sheetData[0]?.length ?? 0;
    if (!isQuickEdit && columnCount > 50) {
      throw new BadRequestException('Too many columns (max 50)');
    }

    const previousCount = request.context?.previousMessages?.length ?? 0;
    if (previousCount > 100) {
      throw new BadRequestException('Conversation history too long');
    }
  }

  private resolveDeclaredRowCount(request: ConversationRequestDto): number {
    const fromCompression = request.sheetCompression?.originalRowCount ?? 0;
    const richContext = request.workbookContext as
      | { sheets?: Array<{ rowCount?: number; sheetName?: string; compressionMeta?: { originalRowCount?: number } }> }
      | undefined;

    const activeSheet =
      richContext && 'activeSheet' in richContext
        ? String((richContext as { activeSheet?: string }).activeSheet ?? '')
        : '';

    const activeSnapshot = richContext?.sheets?.find(
      (sheet) =>
        typeof sheet === 'object' &&
        sheet !== null &&
        (!activeSheet || sheet.sheetName === activeSheet),
    );

    const fromSnapshot = Math.max(
      activeSnapshot?.rowCount ?? 0,
      activeSnapshot?.compressionMeta?.originalRowCount ?? 0,
    );

    return Math.max(fromCompression, fromSnapshot, request.sheetData.length);
  }

  private async getOrCreateConversation(
    conversationId?: string,
    workbookId?: string,
    userId?: string,
  ): Promise<ConversationDocument> {
    if (conversationId) {
      const existing = await this.conversationModel.findOne({ conversationId });
      if (!existing) {
        throw new NotFoundException('CONVERSATION_NOT_FOUND');
      }
      if (existing.expiresAt && existing.expiresAt.getTime() < Date.now()) {
        throw new GoneException('CONVERSATION_EXPIRED');
      }
      // Ownership check (TASKS.md #171). A conversation that already has an
      // owner can only be continued by that owner — otherwise a guessed or
      // leaked conversationId would let one user append to, and read back,
      // another's thread. Reported as NOT_FOUND rather than FORBIDDEN so the
      // response doesn't confirm the id exists.
      if (userId && existing.userId && existing.userId !== userId) {
        throw new NotFoundException('CONVERSATION_NOT_FOUND');
      }
      if (existing.messages.length >= MAX_MESSAGES) {
        throw new BadRequestException('CONTEXT_TOO_LARGE');
      }
      // Backfill only — never overwrite an already-recorded workbookId (mirrors
      // the frontend's own mint-once discipline from TASKS.md #21). Covers a
      // conversation that started before the client had minted/persisted one.
      let dirty = false;
      if (workbookId && !existing.workbookId) {
        existing.workbookId = workbookId;
        dirty = true;
      }
      // Same backfill discipline for userId: claims a pre-#170 conversation for
      // the user continuing it, but never reassigns one that already has an
      // owner (that case threw above).
      if (userId && !existing.userId) {
        existing.userId = userId;
        dirty = true;
      }
      if (dirty) {
        await existing.save();
      }
      return existing;
    }

    const newId = `conv_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
    return this.conversationModel.create({
      conversationId: newId,
      messages: [],
      status: 'active',
      expiresAt: new Date(Date.now() + CONVERSATION_TTL_MS),
      ...(workbookId ? { workbookId } : {}),
      ...(userId ? { userId } : {}),
    });
  }

  /**
   * Full conversation body, scoped to its owner (TASKS.md #171).
   *
   * `userId` is the *caller's* id, resolved from the session by the controller —
   * requesting a conversation owned by someone else is rejected as NOT_FOUND,
   * not silently returned. Unowned (pre-#170) conversations stay readable by id,
   * which is exactly the access level they had before this change.
   */
  async getConversation(conversationId: string, userId?: string) {
    const doc = await this.conversationModel.findOne({ conversationId }).lean();
    if (!doc) {
      throw new NotFoundException('CONVERSATION_NOT_FOUND');
    }
    if (userId && doc.userId && doc.userId !== userId) {
      throw new NotFoundException('CONVERSATION_NOT_FOUND');
    }
    if (doc.expiresAt && doc.expiresAt.getTime() < Date.now()) {
      throw new GoneException('CONVERSATION_EXPIRED');
    }
    return {
      conversationId: doc.conversationId,
      messages: doc.messages ?? [],
      status: doc.status,
      title: doc.title ?? deriveConversationTitle(doc.messages ?? []),
      workbookId: doc.workbookId,
      sheetSnapshot: doc.sheetSnapshot,
      updatedAt: (doc as { updatedAt?: Date }).updatedAt ?? doc.expiresAt,
    };
  }

  /**
   * A user's past conversations, newest first (TASKS.md #171).
   *
   * Returns summaries only — id, title, previews, counts — never message bodies:
   * the list has to stay small enough to load on panel open, and full content is
   * one `getConversation` call away once the user picks one.
   *
   * Pagination is cursor-based on `updatedAt` rather than skip/limit, so a
   * conversation being updated mid-scroll can't shift rows across page
   * boundaries and cause a duplicate or a skip.
   */
  async listConversations(
    userId: string,
    options: { limit?: number; cursor?: string; workbookId?: string } = {},
  ): Promise<{ conversations: ConversationSummary[]; nextCursor: string | null }> {
    const limit = Math.min(
      Math.max(Math.trunc(options.limit ?? HISTORY_DEFAULT_LIMIT), 1),
      HISTORY_MAX_LIMIT,
    );

    const filter: Record<string, unknown> = { userId };
    // #173 is still open (global vs per-workbook history). The list is global by
    // default — the ChatGPT/Cursor model the request was modelled on — but the
    // filter is accepted now so answering #173 the other way is a caller change,
    // not a schema or query rewrite.
    if (options.workbookId) {
      filter.workbookId = options.workbookId;
    }
    if (options.cursor) {
      const cursorDate = new Date(options.cursor);
      if (Number.isNaN(cursorDate.getTime())) {
        throw new BadRequestException('INVALID_CURSOR');
      }
      filter.updatedAt = { $lt: cursorDate };
    }

    // limit + 1 so "is there another page" is answered by the query itself
    // rather than by a second count() that could disagree with it.
    const docs = await this.conversationModel
      .find(filter)
      .sort({ updatedAt: -1 })
      .limit(limit + 1)
      .select('conversationId workbookId title messages status updatedAt createdAt')
      .lean();

    const hasMore = docs.length > limit;
    const page = hasMore ? docs.slice(0, limit) : docs;

    const conversations = page.map((doc) => {
      const messages = doc.messages ?? [];
      const updatedAt = (doc as { updatedAt?: Date }).updatedAt;
      return {
        conversationId: doc.conversationId,
        workbookId: doc.workbookId,
        title: doc.title ?? deriveConversationTitle(messages),
        firstMessage: truncateTitle(messages.find((m) => m.role === 'user')?.content ?? ''),
        lastMessage: truncateTitle(messages[messages.length - 1]?.content ?? ''),
        messageCount: messages.length,
        status: doc.status,
        updatedAt: updatedAt ?? (doc as { createdAt?: Date }).createdAt ?? null,
      };
    });

    const last = page[page.length - 1] as { updatedAt?: Date } | undefined;
    return {
      conversations,
      nextCursor: hasMore && last?.updatedAt ? last.updatedAt.toISOString() : null,
    };
  }

  /**
   * User-set rename, overriding the auto-derived first-message title
   * (TASKS.md #177). Same ownership discipline as `getConversation`/
   * `getOrCreateConversation`: a mismatched owner is reported as NOT_FOUND, not
   * FORBIDDEN, so the response doesn't confirm the id exists. An unowned
   * (pre-#170) conversation may still be renamed by anyone holding its id —
   * the same access level `getConversation` already grants it for reads.
   */
  async renameConversation(
    conversationId: string,
    userId: string,
    title: string,
  ): Promise<{ conversationId: string; title: string }> {
    const trimmed = truncateTitle(title);
    if (!trimmed) {
      throw new BadRequestException('TITLE_REQUIRED');
    }

    const doc = await this.conversationModel.findOne({ conversationId });
    if (!doc) {
      throw new NotFoundException('CONVERSATION_NOT_FOUND');
    }
    if (doc.userId && doc.userId !== userId) {
      throw new NotFoundException('CONVERSATION_NOT_FOUND');
    }

    doc.title = trimmed;
    await doc.save();
    return { conversationId: doc.conversationId, title: trimmed };
  }

  /**
   * Hard delete (TASKS.md #177) — chat history is explicitly framed as
   * read/reopen/rename/delete, not soft-archive, matching the ChatGPT/Cursor
   * baseline this feature was modelled on. There is nothing downstream that
   * references a conversation by its Mongo `_id` in a way a delete would
   * orphan: `change_sets`/`workflow_traces` correlate by `conversationId`
   * string and already tolerate that id resolving to nothing once a
   * conversation expires via TTL (`DATABASE_SCHEMA.md` §4) — a user delete is
   * the same shape of dangling reference, just user-triggered instead of
   * time-triggered.
   */
  async deleteConversation(conversationId: string, userId: string): Promise<void> {
    const doc = await this.conversationModel.findOne({ conversationId }).select('userId').lean();
    if (!doc) {
      throw new NotFoundException('CONVERSATION_NOT_FOUND');
    }
    if (doc.userId && doc.userId !== userId) {
      throw new NotFoundException('CONVERSATION_NOT_FOUND');
    }

    await this.conversationModel.deleteOne({ conversationId });
  }

  private conversationExpiresAt(): Date {
    return new Date(Date.now() + CONVERSATION_TTL_MS);
  }

  private async saveMessage(
    conversationId: string,
    message: ConversationMessageEntry,
  ): Promise<void> {
    // Title is set from the first user message and never rewritten afterwards
    // (TASKS.md #171) — `$setOnInsert` doesn't apply here since the doc already
    // exists, so the "only if absent" condition lives in the filter instead.
    const titleUpdate =
      message.role === 'user' && message.content?.trim()
        ? { title: truncateTitle(message.content) }
        : {};

    await this.conversationModel.updateOne(
      { conversationId },
      {
        $push: { messages: message },
        $set: { updatedAt: new Date(), expiresAt: this.conversationExpiresAt() },
      },
    );

    if (Object.keys(titleUpdate).length > 0) {
      await this.conversationModel.updateOne(
        {
          conversationId,
          $or: [{ title: { $exists: false } }, { title: null }, { title: '' }],
        },
        { $set: titleUpdate },
      );
    }
  }

  private async getRecentMessages(conversationId: string): Promise<ConversationMessageEntry[]> {
    const doc = await this.conversationModel.findOne({ conversationId }).lean();
    return doc?.messages?.slice(-MAX_MESSAGES) ?? [];
  }

  /**
   * CHITCHAT route (see LlmRouterService.classifyIntent): one LOW-tier call,
   * streamed as plain `chunk` SSE events — never `actions`, since there is
   * nothing to preview/accept. Tagged `route: 'chitchat'` in workflow_traces
   * so tier-a-metrics.util.ts (which assumes every traced request could have
   * written something) doesn't fold this in with write-route requests.
   */
  private async handleChitchat(
    activeRequest: ConversationRequestDto,
    request: ConversationRequestDto,
    conversation: ConversationDocument,
    reply: FastifyReply,
    traceId: string,
  ): Promise<void> {
    initSseResponse(reply);
    const conversationId = conversation.conversationId;
    const emit = (event: string, data: Record<string, unknown>) =>
      writeSseEvent(reply, event, { ...data, conversationId });

    this.startWorkflowTrace({
      traceId,
      conversationId,
      workbookId: conversation.workbookId,
      message: activeRequest.message,
      mode: activeRequest.mode,
      request: activeRequest,
    });
    this.workflowTrace.setMeta(traceId, { route: 'chitchat' });

    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}`,
      role: 'user',
      content: request.message,
      type: 'command',
      timestamp: new Date(),
    });

    let fullText = '';
    try {
      for await (const token of this.chitchat.streamReply(activeRequest.message)) {
        fullText += token;
        emit('chunk', { text: token });
      }
    } catch (err) {
      this.logger.error(`Chitchat reply failed trace=${traceId} conversation=${conversationId}`, err);
      fullText = fullText.trim() || "Hi! I'm having trouble responding right now — try again in a moment.";
      emit('chunk', { text: fullText });
    }

    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: fullText.trim(),
      type: 'answer',
      timestamp: new Date(),
    });

    this.workflowTrace.appendNode(traceId, {
      id: 'router',
      type: 'router',
      label: 'Router → chitchat',
      status: 'success',
      output: { route: 'chitchat', responseLength: fullText.length },
      meta: { route: 'chitchat' },
    });

    emit('conversation_end', { summary: 'Ready for your next message.' });
    await this.markCompleted(conversationId);
    this.finalizeWorkflow(traceId, 'completed', {
      route: 'chitchat',
      sseOutput: { kind: 'chitchat', responseLength: fullText.length },
    });

    endSseResponse(reply);
  }

  private async markCompleted(conversationId: string): Promise<void> {
    await this.conversationModel.updateOne(
      { conversationId },
      { $set: { status: 'completed', updatedAt: new Date() } },
    );
  }

  private buildRouterInput(
    request: ConversationRequestDto,
    recentHistory: string[],
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
  ): RouterInput {
    const activeSheet = this.resolveActiveSheetName(request);
    const workbookCtx = request.workbookContext as
      | { activeSheet?: string; sheets?: Array<Record<string, unknown>> }
      | undefined;
    const sheets = workbookCtx?.sheets ?? [];
    const activeSheetData =
      sheets.find(
        (sheet) => sheet.name === activeSheet || sheet.sheetName === activeSheet,
      ) ?? sheets[0];
    const headerRow =
      (activeSheetData?.headers as string[] | undefined) ??
      ((activeSheetData?.rows as unknown[][] | undefined)?.[0] as string[] | undefined);
    const headers = headerRow?.length ? headerRow.map(String) : analysis.headers;

    return {
      message: request.message,
      mode: normalizeAssistantMode(request.mode),
      sheetHeaders: headers,
      activeSheet,
      recentHistory,
    };
  }

  private applyRoutedPromptContext(
    request: ConversationRequestDto,
    conversationId: string,
    decision: RouterDecision,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    traceId: string,
  ): ConversationRequestDto {
    if (decision.route === 'shortcut') {
      return request;
    }

    const rawToon = request.promptContext ?? '';
    const cachedContext = this.contextCache.get(conversationId, rawToon);
    if (cachedContext) {
      this.logger.debug(`[${traceId}] Using cached promptContext`);
      return { ...request, promptContext: cachedContext };
    }

    const activeSheet = this.resolveActiveSheetName(request);
    const tiered = buildTieredToon({
      route: decision.route,
      workbookContext: request.workbookContext ?? {
        activeSheet,
        sheets: [
          {
            name: activeSheet,
            headers: analysis.headers,
            rows: request.sheetData,
          },
        ],
      },
      rawToonPayload: rawToon || undefined,
    });

    if (tiered.promptContext) {
      this.contextCache.set(conversationId, rawToon, tiered.promptContext);
    }

    return {
      ...request,
      promptContext: tiered.promptContext || request.promptContext,
    };
  }

  private async handleRouterShortcut(
    request: ConversationRequestDto,
    routerDecision: RouterDecision,
    conversationId: string,
    traceId: string,
    reply: FastifyReply,
    history: ConversationMessageEntry[],
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<void> {
    const activeSheetName = this.resolveActiveSheetName(request);
    const shortcutActions = routeShortcutAction(request.message, activeSheetName);

    if (!shortcutActions?.length) {
      this.logger.warn(`[${traceId}] Shortcut router returned null — falling back to write`);
      await this.streamWithOrchestrator(
        request,
        reply,
        conversationId,
        traceId,
        history,
        analysis,
        emit,
      );
      return;
    }

    this.logger.log(
      `Shortcut action routed trace=${traceId} conversation=${conversationId} actions=${shortcutActions.map((action) => action.type).join(',')}`,
    );
    await this.emitLocalDecision(
      conversationId,
      {
        kind: 'actions',
        answer: buildShortcutAnswer(shortcutActions),
        explanation: routerDecision.reasoning || 'Routed via LLM shortcut handler.',
        actions: shortcutActions,
      },
      emit,
      { traceId, route: 'shortcut', tier: 0 },
    );
    endSseResponse(reply);
  }

  /**
   * Persist LLM call telemetry into the workflow trace (task #92).
   *
   * Provider/model/token counts previously existed only in `logger.log` output,
   * which no file logger captures — so any LLM failure became undiagnosable the
   * moment the terminal scrolled. Reasoning-token exhaustion (#86) is called out
   * explicitly because it is invisible in a plain token count: completion tokens
   * are spent while the returned content is empty.
   */
  private recordLlmCallTrace(
    traceId: string,
    telemetry: LlmCallTelemetry,
    outcome: { durationMs: number; emptyResponse: boolean; structuredKind: string },
  ): void {
    const usage = telemetry.usage;
    const completionTokens = usage?.completionTokens ?? 0;
    const reasoningExhausted = outcome.emptyResponse && completionTokens > 0;
    this.workflowTrace.appendNode(traceId, {
      id: `llm_${Date.now()}`,
      type: outcome.emptyResponse ? 'error' : 'sse_out',
      label: reasoningExhausted
        ? 'LLM returned no content (reasoning-token exhaustion)'
        : 'LLM call',
      status: outcome.emptyResponse ? 'failed' : 'success',
      durationMs: outcome.durationMs,
      meta: {
        provider: telemetry.provider ?? 'unknown',
        model: telemetry.model ?? 'unknown',
        modelTier: telemetry.modelTier ?? 'unknown',
        promptTokens: usage?.promptTokens ?? null,
        completionTokens: usage?.completionTokens ?? null,
        totalTokens: usage?.totalTokens ?? null,
        reasoningTokens: usage?.reasoningTokens ?? null,
        emptyResponse: outcome.emptyResponse,
        reasoningExhausted,
        structuredKind: outcome.structuredKind,
      },
    });
  }

  private formatUsage(telemetry: LlmCallTelemetry): string {
    const usage = telemetry.usage;
    if (!usage) {
      return 'unavailable';
    }

    const prompt = usage.promptTokens ?? '-';
    const completion = usage.completionTokens ?? '-';
    const total = usage.totalTokens ?? '-';
    return `prompt:${prompt},completion:${completion},total:${total}`;
  }

  private clipForLog(value: string, maxLength = 500): string {
    const normalized = value.replace(/\s+/g, ' ').trim();
    if (normalized.length <= maxLength) {
      return normalized;
    }
    return `${normalized.slice(0, maxLength)}...`;
  }
}
