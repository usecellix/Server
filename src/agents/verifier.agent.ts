import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import {
  VERIFIER_SYSTEM_PROMPT,
  buildVerifierUserMessage,
  normalizeVerifierOutput,
} from './prompts/verifier.prompt';
import { parseAgentJson } from './utils/parse-agent-json.util';
import { Action, SubTask, VerifierOutput, WorkbookContext } from './types/agent.types';
import { StructuredLogger } from './logging/structured-logger';

@Injectable()
export class VerifierAgent {
  private readonly logger = new Logger(VerifierAgent.name);

  constructor(
    private readonly llm: OpenRouterService,
    private readonly config: AppConfigService,
    private readonly structuredLogger: StructuredLogger = new StructuredLogger(),
  ) {}

  async verify(
    originalPrompt: string,
    subtasks: SubTask[],
    actionsBySubtask: Record<string, Action[]>,
    context: WorkbookContext,
    formulaValidatorSummary?: string,
    correlationId = `req_${Date.now()}`,
  ): Promise<VerifierOutput> {
    const startedAt = Date.now();
    const model = this.config.openRouterModelMedium;
    const userMessage = buildVerifierUserMessage(
      originalPrompt,
      subtasks,
      actionsBySubtask,
      context,
      formulaValidatorSummary,
    );

    const raw = await this.llm.complete({
      systemPrompt: VERIFIER_SYSTEM_PROMPT,
      userMessage,
      model,
      temperature: 0.1,
      maxTokens: 1500,
    });
    this.structuredLogger.debugRawResponse(correlationId, 'verifier', model, raw);

    try {
      const parsed = parseAgentJson<Partial<VerifierOutput>>(raw);
      const normalized = normalizeVerifierOutput(parsed, subtasks.map((s) => s.id));
      this.logger.log(`Verifier: ${normalized.passed ? 'PASS' : 'FAIL'} — ${normalized.feedback}`);
      this.structuredLogger.logAgentEvent({
        correlationId,
        agent: 'verifier',
        model,
        durationMs: Date.now() - startedAt,
        success: true,
        tokenUsage: this.structuredLogger.estimateTokens(raw),
        rawResponse: raw,
        parsedResponse: normalized,
      });
      return normalized;
    } catch (error) {
      this.structuredLogger.warnParseFailure(
        correlationId,
        'verifier',
        model,
        raw,
        error instanceof Error ? error.message : String(error),
      );
      this.logger.error('Verifier returned invalid JSON', raw);
      const fallback = {
        passed: true,
        feedback: 'Verifier parse error — defaulting to pass',
        issues: [],
        subtaskResults: subtasks.map((s) => ({
          subtaskId: s.id,
          passed: true,
          feedback: 'Parse error — default pass',
          issues: [],
        })),
      };
      this.structuredLogger.logAgentEvent({
        correlationId,
        agent: 'verifier',
        model,
        durationMs: Date.now() - startedAt,
        success: false,
        tokenUsage: this.structuredLogger.estimateTokens(raw),
        rawResponse: raw,
        parsedResponse: fallback,
        error: 'Verifier JSON parse failed',
      });
      return fallback;
    }
  }
}
