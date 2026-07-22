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
import { classifyIntent, detectAmbiguity } from '../llm/ambiguity-detector';
import { LLMTier, SheetSnapshot } from '../../types/cellix.types';
import { OrchestratorService } from '../../agents/orchestrator.service';
import { SseEmitter } from '../../agents/sse.emitter';
import { ToolBridgeService } from '../../agents/tool-bridge.service';
import { buildAgentWorkbookContext } from '../../agents/utils/workbook-context.builder';
import { WriteRouteNoActionError } from '../errors/write-route-no-action.error';
import { ConversationEngineService, EngineResponse, LlmRequestError } from './conversation-engine.service';
import { DataQueryService } from './data-query.service';
import { FindExportService, FindExportSheetSlice } from './find-export.service';
import { ContextCacheService } from './context-cache.service';
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
import {
  buildTableActionsFromMessage,
  detectCreateNewSheetIntent,
  detectSheetDataGenerationIntent,
  parseTableCreateRequest,
} from '../utils/table-request.util';
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
import { buildEnrichedPromptContext } from '../../formula/enrich-context.util';
import { FormulaAnalyzer } from '../../formula/formula.analyzer';
import { SmartDataQueryService } from './smart-data-query.service';
import { SheetAnalyzerService } from './sheet-analyzer.service';
import { Tier0DirectService, Tier0Result } from './tier0-direct.service';
import { Tier1SingleActionService } from './tier1-single-action.service';
import { Tier2GenerateVerifyService } from './tier2-generate-verify.service';
import { StructuredLogger } from '../../agents/logging/structured-logger';
import { WorkbookContext as AgentWorkbookContext } from '../../agents/types/agent.types';
import { SheetAction } from '../types/sheet-actions.types';
import { isFindLookupMessage } from '../utils/find-query-parser.util';
import {
  buildInternalDetails,
  buildUserFacingSummary,
  tierProcessingLabel,
} from '../utils/user-facing-response.util';

const MAX_MESSAGES = 50;
const CONVERSATION_TTL_MS = Number(process.env.CONVERSATION_TTL_HOURS ?? 168) * 60 * 60 * 1000;

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
  ) {}

  private enrichAgentContext(
    context: AgentWorkbookContext,
    basePromptContext?: string,
    history?: ConversationMessageEntry[],
    userMessage?: string,
  ): { enrichedContext: AgentWorkbookContext; promptContext: string } {
    const enrichedSheets = context.sheets.map((sheet) => ({
      ...sheet,
      formulaInsights: this.formulaAnalyzer.analyzeSheet(sheet),
    }));
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

  async handleConversation(
    request: ConversationRequestDto,
    reply: FastifyReply,
    traceId = '-',
  ): Promise<void> {
    this.validateRequest(request);

    const conversation = await this.getOrCreateConversation(request.conversationId);
    const activeRequestRaw = await this.applyRefinementContext(request);
    const activeRequest: ConversationRequestDto = {
      ...activeRequestRaw,
      mode: normalizeAssistantMode(activeRequestRaw.mode),
    };
    const requestMode = activeRequest.mode ?? 'action';
    const writeAllowed = requestMode === 'action';

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
        );
        endSseResponse(reply);
        return;
      }
      // Regex matched but handler returned null — fall through to full path.
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

      const tablePlan = parseTableCreateRequest(request.message);
      const tableActions = buildTableActionsFromMessage(request.message);
      const isNewSheetWithData =
        detectCreateNewSheetIntent(request.message) &&
        detectSheetDataGenerationIntent(request.message);
      if (
        writeAllowed &&
        tablePlan &&
        tableActions?.length &&
        tablePlan.headers.length >= 2 &&
        !isNewSheetWithData
      ) {
        this.logger.log(
          `Table create (deterministic) trace=${traceId} conversation=${conversationId} rows=${tablePlan.rowCount} cols=${tablePlan.headers.length}`,
        );
        const decision = {
          kind: 'actions' as const,
          answer: `Created **${tablePlan.rowCount}** rows with columns: ${tablePlan.headers.join(', ')}.`,
          explanation: 'Wrote headers and all data rows to your sheet.',
          actions: tableActions,
        };
        await this.emitLocalDecision(conversation.conversationId, decision, emit);
        endSseResponse(reply);
        return;
      }

      const recentHistory = history
        .filter((entry) => entry.role === 'user')
        .slice(-2)
        .map((entry) => entry.content);

      const routerDecision = await this.llmRouter.route(
        this.buildRouterInput(activeRequest, recentHistory, analysis),
      );

      this.logger.log(
        `[${traceId}] Router: route=${routerDecision.route} confidence=${routerDecision.confidence} "${routerDecision.reasoning}"`,
      );

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
          await this.emitLocalDecision(conversationId, exportDecision, emit);
          endSseResponse(reply);
          return;
        }
      }

      if (this.engine.hasOpenAi()) {
        if (routerDecision.route === 'ask') {
          const ambiguityOutcome = await this.checkAmbiguity(routedRequest, analysis, history);
          if (ambiguityOutcome?.clarification) {
            await this.emitClarification(
              conversationId,
              ambiguityOutcome.clarification,
              emit,
              reply,
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
          this.logger.warn(`LLM unavailable, using local engine: ${reason}`);
          emit('status', { message: 'AI unavailable — limited local mode…' });
        }
      } else {
        emit('status', { message: 'AI not configured — set OPENROUTER_API_KEY in backend .env' });
      }

      const decision = this.engine.decide(
        activeRequest.message,
        activeRequest.sheetData,
        analysis,
        history,
        resolveEngineWorkbookMeta(activeRequest),
      );
      this.logger.log(
        `AI skipped trace=${traceId} conversation=${conversationId} provider=local reason=${localReason} result=${decision.kind} durationMs=${Date.now() - startedAt}`,
      );
      await this.emitLocalDecision(conversation.conversationId, decision, emit);
      endSseResponse(reply);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to process your request';
      this.logger.error(message, error instanceof Error ? error.stack : undefined);
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
      return;
    }

    if (decision.kind === 'actions') {
      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: decision.answer,
        type: 'answer',
        timestamp: new Date(),
        metadata: this.buildWriteMetadata(decision.actions),
      });

      emit('answer', { answer: decision.answer });
      emit('actions', {
        actions: decision.actions,
        explanation: decision.explanation,
      });
      emit('conversation_end', { summary: 'Changes applied.' });
      await this.markCompleted(conversationId);
      return;
    }

    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: decision.answer,
      type: 'answer',
      timestamp: new Date(),
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
        );
        endSseResponse(reply);
        return;
      }

      const classifiedTier = (routerDecision.complexity ?? 3) as 0 | 1 | 2 | 3;
      const tieringMode = getComplexityTieringMode();
      const complexity = resolveExecutableTier(classifiedTier, tieringMode);
      const actionHint = routerDecision.actionHint;
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
            if (tier1Result.actions.length > 0) {
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
        );
        return;
      }

      outcome.tier = 3;
      outcome.llmCallCount = 3;
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
      });
    }
  }

  private resolveInitialWriteOutcome(routerDecision: RouterDecision): {
    tier: 0 | 1 | 2 | 3;
    llmCallCount: number;
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
      endSseResponse(reply);
      return;
    }

    const actions = this.engine.finalizeActions(
      result.actions,
      analysis,
      richWorkbookContext,
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
      tier: 2,
      durationMs: result.durationMs,
    });
    emit('conversation_end', { summary: 'Review changes and accept or reject.', tier: 2 });
    await this.markCompleted(conversationId);
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
      answer,
      processingLabel,
      actions,
      emit,
      traceId,
      prompt,
      agentContext,
      assumption,
      model,
    } = params;

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
      tier,
    });
    emit('conversation_end', { summary: 'Review changes and accept or reject.', tier });
    await this.markCompleted(conversationId);
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
        },
        sseEmitter,
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
          const failedReason =
            orchestratorResult.failedSubtask?.reason ??
            'A later step could not be completed';
          const actions = this.engine.finalizeActions(rawActions, analysis, richWorkbookContext);
          actionsCount = actions.length;

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
            partialProgress: true,
            failedSubtask: orchestratorResult.failedSubtask,
            tier: 3,
          });
          emit('conversation_end', {
            summary: 'Partial changes ready — review and accept, or retry the failed step.',
          });
          await this.markCompleted(conversationId);
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

      const actions = this.engine.finalizeActions(rawActions, analysis, richWorkbookContext);
      actionsCount = actions.length;

      this.assertWriteRouteProducedActions({
        conversationId,
        message: request.message,
        actionsLength: actions.length,
      });

      const answer = `I'll apply the prepared changes to your sheet.`;

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

      const processingLabel = tierProcessingLabel(3);
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
        metadata: this.buildWriteMetadata(actions, changeSet.changeSetId),
      });

      emit('answer', { answer, tier: 3 });
      emit('actions', {
        actions,
        explanation: processingLabel,
        userFacingSummary,
        internalDetails,
        changeSetId: changeSet.changeSetId,
        changes: changeSet.changes,
        tier: 3,
      });
      emit('conversation_end', { summary: 'Review changes and accept or reject.', tier: 3 });
      await this.markCompleted(conversationId);
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
      const plan = await this.orchestrator.planOnly({
        prompt: request.message,
        context: enrichedContext,
        conversationHistory,
        promptContext,
        conversationId,
        correlationId: traceId,
        complexity,
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
      const fallbackText = fullText.trim() || 'I could not generate a response.';
      if (structured?.kind === 'actions') {
        actionsCount = structured.actions.length;
      }
      this.logger.log(
        `AI response trace=${traceId} conversation=${conversationId} called=true provider=${telemetry.provider ?? 'unknown'} modelTier=${telemetry.modelTier ?? 'unknown'} model=${telemetry.model ?? 'unknown'} tokens=${this.formatUsage(telemetry)} durationMs=${Date.now() - startedAt} response="${this.clipForLog(fallbackText)}"`,
      );

      if (!structured) {
        const isNewSheetWithData =
          detectCreateNewSheetIntent(request.message) &&
          detectSheetDataGenerationIntent(request.message);
        const tableFallback =
          readOnly || isNewSheetWithData
            ? null
            : buildTableActionsFromMessage(request.message);
        if (tableFallback?.length) {
          const plan = parseTableCreateRequest(request.message);
          const answer = plan
            ? `Created **${plan.rowCount}** rows with columns: ${plan.headers.join(', ')}.`
            : 'Created your table with headers and sample values.';
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
          endSseResponse(reply);
          return;
        }

        const retryHint =
          'I understood your request but could not parse the AI response. Please try again — e.g. "Generate 10 rows of sample GST purchase data with headers".';
        const answer =
          fallbackText.length > 20 && !fallbackText.startsWith('{')
            ? `${fallbackText}\n\n${retryHint}`
            : retryHint;

        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: answer,
          type: 'answer',
          timestamp: new Date(),
        });
        emit('answer', { answer });
        emit('conversation_end', { summary: 'Completed.' });
        await this.markCompleted(conversationId);
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
            ? `\n\n_${modeLabel} mode is read-only. Switch to Action mode to apply these changes._`
            : '';
        const answer = `${structured.answer}${note}`;
        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: answer,
          type: 'answer',
          timestamp: new Date(),
        });
        emit('answer', { answer });
        emit('conversation_end', { summary: 'Read-only response.' });
        await this.markCompleted(conversationId);
        endSseResponse(reply);
        return;
      }

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
      endSseResponse(reply);
      return;
    }

    await this.saveMessage(conversationId, {
      id: `msg_${Date.now()}_assistant`,
      role: 'assistant',
      content: structured.answer,
      type: 'answer',
      timestamp: new Date(),
    });
    emit('answer', { answer: structured.answer });
    emit('conversation_end', { summary: 'Completed.' });
    await this.markCompleted(conversationId);
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

  private async getOrCreateConversation(conversationId?: string): Promise<ConversationDocument> {
    if (conversationId) {
      const existing = await this.conversationModel.findOne({ conversationId });
      if (!existing) {
        throw new NotFoundException('CONVERSATION_NOT_FOUND');
      }
      if (existing.expiresAt && existing.expiresAt.getTime() < Date.now()) {
        throw new GoneException('CONVERSATION_EXPIRED');
      }
      if (existing.messages.length >= MAX_MESSAGES) {
        throw new BadRequestException('CONTEXT_TOO_LARGE');
      }
      return existing;
    }

    const newId = `conv_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
    return this.conversationModel.create({
      conversationId: newId,
      messages: [],
      status: 'active',
      expiresAt: new Date(Date.now() + CONVERSATION_TTL_MS),
    });
  }

  async getConversation(conversationId: string) {
    const doc = await this.conversationModel.findOne({ conversationId }).lean();
    if (!doc) {
      throw new NotFoundException('CONVERSATION_NOT_FOUND');
    }
    if (doc.expiresAt && doc.expiresAt.getTime() < Date.now()) {
      throw new GoneException('CONVERSATION_EXPIRED');
    }
    return {
      conversationId: doc.conversationId,
      messages: doc.messages ?? [],
      status: doc.status,
      sheetSnapshot: doc.sheetSnapshot,
      updatedAt: (doc as { updatedAt?: Date }).updatedAt ?? doc.expiresAt,
    };
  }

  private conversationExpiresAt(): Date {
    return new Date(Date.now() + CONVERSATION_TTL_MS);
  }

  private async saveMessage(
    conversationId: string,
    message: ConversationMessageEntry,
  ): Promise<void> {
    await this.conversationModel.updateOne(
      { conversationId },
      {
        $push: { messages: message },
        $set: { updatedAt: new Date(), expiresAt: this.conversationExpiresAt() },
      },
    );
  }

  private async getRecentMessages(conversationId: string): Promise<ConversationMessageEntry[]> {
    const doc = await this.conversationModel.findOne({ conversationId }).lean();
    return doc?.messages?.slice(-MAX_MESSAGES) ?? [];
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
    );
    endSseResponse(reply);
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
