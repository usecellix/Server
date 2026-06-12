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
import { classifyIntent, detectAmbiguity } from '../llm/ambiguity-detector';
import { LLMTier } from '../../types/cellix.types';
import { ConversationEngineService, EngineResponse, LlmRequestError } from './conversation-engine.service';
import { LlmCallTelemetry, OpenRouterService } from './openrouter.service';
import { buildStatusMessage } from '../utils/status-message.util';
import { buildTableActionsFromMessage, parseTableCreateRequest } from '../utils/table-request.util';
import {
  resolveConversationHistory,
  resolveEngineWorkbookMeta,
  resolveWorkbookContext,
} from '../utils/workbook-context-resolver.util';
import { SheetAnalyzerService } from './sheet-analyzer.service';

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
    private readonly openRouter: OpenRouterService,
  ) {}

  async handleConversation(
    request: ConversationRequestDto,
    reply: FastifyReply,
    traceId = '-',
  ): Promise<void> {
    this.validateRequest(request);

    const conversation = await this.getOrCreateConversation(request.conversationId);
    const analysis = this.sheetAnalyzer.analyze(request.sheetData);
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
        `Conversation request trace=${traceId} conversation=${conversationId} message="${this.clipForLog(request.message)}" sheet=${analysis.rowCount}x${analysis.columnCount} history=${request.context?.previousMessages?.length ?? 0}`,
      );
      emit('status', { message: buildStatusMessage(request.message, analysis) });

      const history = await this.getRecentMessages(conversation.conversationId);

      const tablePlan = parseTableCreateRequest(request.message);
      const tableActions = buildTableActionsFromMessage(request.message);
      if (tablePlan && tableActions?.length && tablePlan.headers.length >= 2) {
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
        const ambiguityOutcome = await this.checkAmbiguity(request, analysis);
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

        try {
          await this.streamWithOpenAi(request, reply, conversationId, traceId, history, analysis, emit);
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
        request.message,
        request.sheetData,
        analysis,
        history,
        resolveEngineWorkbookMeta(request),
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
    emit('conversation_end', { summary: 'Ready for your next message.' });
    await this.markCompleted(conversationId);
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
        const tableFallback = buildTableActionsFromMessage(request.message);
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

  private validateRequest(request: ConversationRequestDto): void {
    if (!request.message?.trim()) {
      throw new BadRequestException('Message is required');
    }

    if (!Array.isArray(request.sheetData)) {
      throw new BadRequestException('Invalid sheet data format');
    }

    if (request.sheetData.length > 1000) {
      throw new BadRequestException('Sheet too large (max 1000 rows)');
    }

    const columnCount = request.sheetData[0]?.length ?? 0;
    if (columnCount > 50) {
      throw new BadRequestException('Too many columns (max 50)');
    }

    const previousCount = request.context?.previousMessages?.length ?? 0;
    if (previousCount > 100) {
      throw new BadRequestException('Conversation history too long');
    }
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
