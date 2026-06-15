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
import { ChangeSetService } from '../../audit/change-set.service';
import { classifyIntent, detectAmbiguity } from '../llm/ambiguity-detector';
import { LLMTier, SheetSnapshot } from '../../types/cellix.types';
import { OrchestratorService } from '../../agents/orchestrator.service';
import { SseEmitter } from '../../agents/sse.emitter';
import { ToolBridgeService } from '../../agents/tool-bridge.service';
import { buildAgentWorkbookContext } from '../../agents/utils/workbook-context.builder';
import { ConversationEngineService, EngineResponse, LlmRequestError } from './conversation-engine.service';
import { DataQueryService, FindMatch } from './data-query.service';
import { IntentClassifierService, intentIsReadOnly } from './intent-classifier.service';
import { LlmCallTelemetry, OpenRouterService } from './openrouter.service';
import { AssistantMode, IntentType } from '../types/sheet-actions.types';
import {
  ASK_MODE_READONLY_DIRECTIVE,
  PLAN_MODE_DIRECTIVE,
} from '../prompt/cellix-system-prompt';
import { modeIsReadOnly, stripWriteActions } from '../utils/mode-guard.util';
import { PlannerOutput } from '../../agents/types/agent.types';
import { buildStatusMessage } from '../utils/status-message.util';
import {
  buildTableActionsFromMessage,
  detectCreateNewSheetIntent,
  detectSheetDataGenerationIntent,
  parseTableCreateRequest,
} from '../utils/table-request.util';
import {
  resolveConversationHistory,
  resolveEngineWorkbookMeta,
  resolveWorkbookContext,
} from '../utils/workbook-context-resolver.util';
import { buildRefinementContext } from '../utils/refinement-context.util';
import { buildEnrichedPromptContext } from '../../formula/enrich-context.util';
import { FormulaAnalyzer } from '../../formula/formula.analyzer';
import { SheetAnalyzerService } from './sheet-analyzer.service';
import { WorkbookContext as AgentWorkbookContext } from '../../agents/types/agent.types';

const MAX_MESSAGES = 50;

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
    private readonly intentClassifier: IntentClassifierService,
    private readonly dataQuery: DataQueryService,
    private readonly formulaAnalyzer: FormulaAnalyzer,
    private readonly toolBridge: ToolBridgeService,
  ) {}

  private enrichAgentContext(
    context: AgentWorkbookContext,
    basePromptContext?: string,
  ): { enrichedContext: AgentWorkbookContext; promptContext: string } {
    const enrichedSheets = context.sheets.map((sheet) => ({
      ...sheet,
      formulaInsights: this.formulaAnalyzer.analyzeSheet(sheet),
    }));
    const enrichedContext: AgentWorkbookContext = { ...context, sheets: enrichedSheets };
    const promptContext = buildEnrichedPromptContext(basePromptContext, enrichedSheets);
    return { enrichedContext, promptContext };
  }

  async handleConversation(
    request: ConversationRequestDto,
    reply: FastifyReply,
    traceId = '-',
  ): Promise<void> {
    this.validateRequest(request);

    const conversation = await this.getOrCreateConversation(request.conversationId);
    const activeRequest = await this.applyRefinementContext(request);
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

      const requestMode = activeRequest.mode;
      const writeAllowed = requestMode !== 'ask' && requestMode !== 'plan';

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

      if (this.engine.hasOpenAi()) {
        const classification = this.intentClassifier.classify(request.message);
        const route = this.resolveRoute(activeRequest.mode, classification.intent);

        if (route === 'readonly') {
          const ambiguityOutcome = await this.checkAmbiguity(activeRequest, analysis);
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

          // For find/search/locate queries run the search locally against the
          // full workbook (fetching on demand when compressed) before the LLM.
          const isFind =
            classification.subIntent === 'find' ||
            (activeRequest.mode === 'ask' &&
              this.intentClassifier.isFindLookupIntent(request.message.toLowerCase()));
          if (isFind) {
            const findDecision = await this.handleFindQuery(
              activeRequest,
              analysis,
              conversationId,
              emit,
            );
            if (findDecision) {
              await this.emitLocalDecision(conversationId, findDecision, emit);
              endSseResponse(reply);
              return;
            }
          }
        }

        try {
          if (route === 'orchestrator') {
            await this.streamWithOrchestrator(
              activeRequest,
              reply,
              conversationId,
              traceId,
              history,
              analysis,
              emit,
            );
          } else if (route === 'planner') {
            await this.streamWithPlanner(
              activeRequest,
              reply,
              conversationId,
              traceId,
              analysis,
              emit,
            );
          } else {
            await this.streamWithOpenAi(
              activeRequest,
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
      emit('error', { message });
      endSseResponse(reply);
    }
  }

  /**
   * Mode-first routing. When the add-in sends an explicit mode we honor it
   * (ask = read-only, plan = planner, action = orchestrator). Older clients
   * without a mode keep the previous intent-based behavior.
   */
  private resolveRoute(
    mode: AssistantMode | undefined,
    intent: IntentType,
  ): 'orchestrator' | 'readonly' | 'planner' {
    if (mode === 'action') return 'orchestrator';
    if (mode === 'plan') return 'planner';
    if (mode === 'ask') return 'readonly';
    return intentIsReadOnly(intent) ? 'readonly' : 'orchestrator';
  }

  private async checkAmbiguity(
    request: ConversationRequestDto,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
  ) {
    const workbookContext = resolveWorkbookContext(request, analysis, request.sheetData);
    const conversationHistory = resolveConversationHistory(request);
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

  private async handleFindQuery(
    request: ConversationRequestDto,
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    conversationId: string,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<EngineResponse | null> {
    const term = this.dataQuery.extractSearchTerm(request.message);
    if (!term) {
      return {
        kind: 'answer',
        answer:
          'I could not extract a search value from your message. Please try: "Find CGST value 1868".',
      };
    }

    const activeSheetName = this.resolveActiveSheetName(request);
    emit('status', { message: 'Searching across your workbook…' });

    const richContext = resolveWorkbookContext(request, analysis, request.sheetData);
    const allMatches: FindMatch[] = [];

    // 1) Active sheet — reuse on-demand fetch for compressed payloads.
    const activeData = await this.resolveActiveSheetData(
      request,
      analysis,
      activeSheetName,
      conversationId,
      emit,
    );
    const activeAnalysis = this.sheetAnalyzer.analyze(activeData);
    allMatches.push(
      ...this.dataQuery.collectMatches(request.message, activeData, activeAnalysis, activeSheetName),
    );

    // 2) Every other sheet in the workbook (full cross-sheet awareness).
    for (const snapshot of richContext.sheets ?? []) {
      if (!snapshot?.sheetName || snapshot.sheetName === activeSheetName) continue;
      const data = await this.resolveSnapshotData(snapshot, conversationId, emit);
      if (!data.length) continue;
      const sheetAnalysis = this.sheetAnalyzer.analyze(data);
      allMatches.push(
        ...this.dataQuery.collectMatches(
          request.message,
          data,
          sheetAnalysis,
          snapshot.sheetName,
        ),
      );
    }

    const result = this.dataQuery.buildFindResult(term, allMatches);
    return {
      kind: 'answer',
      answer: result.answer,
      followUp: result.followUp,
      selectCell: result.selectCell,
      matches: result.matches,
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
        metadata: { actions: decision.actions },
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

    emit('answer', { answer: decision.answer });
    if (decision.kind === 'answer' && decision.matches?.length) {
      emit('matches', { matches: decision.matches, summary: decision.answer });
    }
    if (decision.kind === 'answer' && decision.selectCell) {
      emit('select_cell', decision.selectCell);
    }
    emit('conversation_end', { summary: 'Ready for your next message.' });
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
    const { enrichedContext, promptContext } = this.enrichAgentContext(agentContext, basePrompt);

    const conversationHistory = resolveConversationHistory(request).map((entry) => ({
      role: entry.role as 'user' | 'assistant',
      content: entry.content,
    }));

    const sseEmitter = new SseEmitter(emit);

    try {
      const rawActions = await this.orchestrator.run(
        {
          prompt: request.message,
          context: enrichedContext,
          conversationHistory,
          promptContext,
          conversationId,
          toolEmit: emit,
        },
        sseEmitter,
      );

      if (rawActions.length === 0) {
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

      const actions = this.engine.finalizeActions(rawActions, analysis, richWorkbookContext);
      actionsCount = actions.length;

      if (actions.length === 0) {
        const question =
          'I understood your request but could not produce valid actions. Can you clarify what to change?';
        await this.saveMessage(conversationId, {
          id: `msg_${Date.now()}_assistant`,
          role: 'assistant',
          content: question,
          type: 'question',
          timestamp: new Date(),
        });
        emit('question', { question });
        endSseResponse(reply);
        success = true;
        return;
      }

      const answer = `Here are **${actions.length}** change(s) I will apply to your sheet.`;

      const changeSet = await this.changeSetService.createPreview({
        conversationId,
        traceId,
        prompt: request.message,
        context: enrichedContext,
        actions,
      });

      await this.saveMessage(conversationId, {
        id: `msg_${Date.now()}_assistant`,
        role: 'assistant',
        content: answer,
        type: 'answer',
        timestamp: new Date(),
        metadata: { actions, changeSetId: changeSet.changeSetId },
      });

      emit('answer', { answer });
      emit('actions', {
        actions,
        explanation: 'Multi-agent pipeline: planned, executed, and verified.',
        changeSetId: changeSet.changeSetId,
        changes: changeSet.changes,
      });
      emit('conversation_end', { summary: 'Review changes and accept or reject.' });
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
      throw error;
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
    analysis: ReturnType<SheetAnalyzerService['analyze']>,
    emit: (event: string, data: Record<string, unknown>) => void,
  ): Promise<void> {
    const startedAt = Date.now();
    const intent = classifyIntent(request.message);
    const telemetry: LlmCallTelemetry = { provider: 'openrouter', modelTier: 'high' };
    let success = false;

    const richWorkbookContext = resolveWorkbookContext(request, analysis, request.sheetData);
    const agentContext = buildAgentWorkbookContext(richWorkbookContext, request.sheetData, analysis);
    const basePrompt = request.promptContext ?? richWorkbookContext.prompt_context ?? undefined;
    const { enrichedContext, promptContext } = this.enrichAgentContext(agentContext, basePrompt);
    const conversationHistory = resolveConversationHistory(request).map((entry) => ({
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
      throw error;
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
    const conversationTurns = resolveConversationHistory(request);
    const llmPlan = this.engine.planLlmCall(
      request.message,
      request.sheetData,
      analysis,
      history,
      resolveEngineWorkbookMeta(request),
      richWorkbookContext,
      conversationTurns,
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
      expiresAt: new Date(Date.now() + 24 * 60 * 60 * 1000),
    });
  }

  private async saveMessage(
    conversationId: string,
    message: ConversationMessageEntry,
  ): Promise<void> {
    await this.conversationModel.updateOne(
      { conversationId },
      {
        $push: { messages: message },
        $set: { updatedAt: new Date(), expiresAt: new Date(Date.now() + 24 * 60 * 60 * 1000) },
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
