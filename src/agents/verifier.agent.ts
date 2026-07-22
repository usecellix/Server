import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import {
  VERIFIER_SYSTEM_PROMPT,
  buildVerifierUserMessage,
  normalizeVerifierOutput,
} from './prompts/verifier.prompt';
import { parseAgentJson } from './utils/parse-agent-json.util';
import { salvageVerifierSubtaskResults } from './utils/verifier-partial-parse.util';
import {
  VERIFIER_LAST_RESORT_MAX_TOKENS,
  VERIFIER_REASONING_MAX_TOKENS,
  resolveVerifierMaxTokens,
} from './utils/verifier-token-budget.util';
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
    const subtaskIds = subtasks.map((s) => s.id);
    const userMessage = buildVerifierUserMessage(
      originalPrompt,
      subtasks,
      actionsBySubtask,
      context,
      formulaValidatorSummary,
    );

    const maxTokens = resolveVerifierMaxTokens(subtasks.length);
    let raw = await this.llm.complete({
      systemPrompt: VERIFIER_SYSTEM_PROMPT,
      userMessage,
      model,
      temperature: 0.1,
      maxTokens,
      reasoningEffort: 'low',
      reasoningMaxTokens: VERIFIER_REASONING_MAX_TOKENS,
    });
    this.structuredLogger.debugRawResponse(correlationId, 'verifier', model, raw);

    let normalized = this.tryNormalize(raw, subtaskIds, correlationId, model);
    const needsRetry =
      !normalized ||
      normalized.subtaskResults.some((r) => r.inconclusive) ||
      this.looksTruncated(raw, subtaskIds.length);

    if (needsRetry) {
      this.logger.warn(
        `Verifier response incomplete/truncated — retrying verify-only with maxTokens=${VERIFIER_LAST_RESORT_MAX_TOKENS}`,
      );
      raw = await this.llm.complete({
        systemPrompt: VERIFIER_SYSTEM_PROMPT,
        userMessage:
          userMessage +
          '\n\nIMPORTANT: Return COMPLETE JSON for EVERY subtask in subtaskResults. Do not truncate.',
        model,
        temperature: 0.1,
        maxTokens: VERIFIER_LAST_RESORT_MAX_TOKENS,
        reasoningEffort: 'low',
        reasoningMaxTokens: VERIFIER_REASONING_MAX_TOKENS,
      });
      this.structuredLogger.debugRawResponse(correlationId, 'verifier', model, raw);
      const retried = this.tryNormalize(raw, subtaskIds, correlationId, model);
      if (retried) {
        normalized = this.mergePreferringResolved(normalized, retried, subtaskIds);
      }
    }

    if (normalized) {
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
    }

    this.logger.error('Verifier returned invalid JSON after retry', raw);
    // Soft-pass only when we truly cannot parse — do not invent per-subtask failures.
    const fallback: VerifierOutput = {
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

  private tryNormalize(
    raw: string,
    subtaskIds: string[],
    correlationId: string,
    model: string,
  ): VerifierOutput | null {
    try {
      const parsed = parseAgentJson<Partial<VerifierOutput>>(raw);
      return normalizeVerifierOutput(parsed, subtaskIds, {
        fillMissingAsInconclusive: true,
      });
    } catch (error) {
      this.structuredLogger.warnParseFailure(
        correlationId,
        'verifier',
        model,
        raw,
        error instanceof Error ? error.message : String(error),
      );

      const salvaged = salvageVerifierSubtaskResults(raw);
      if (salvaged.length > 0) {
        this.logger.warn(
          `Verifier JSON truncated — salvaged ${salvaged.length}/${subtaskIds.length} subtaskResults`,
        );
        return normalizeVerifierOutput(
          {
            passed: false,
            feedback: 'Partial verifier response (truncated)',
            issues: [],
            subtaskResults: salvaged,
          },
          subtaskIds,
          { fillMissingAsInconclusive: true },
        );
      }

      return null;
    }
  }

  private looksTruncated(raw: string, subtaskCount: number): boolean {
    if (!raw?.trim()) return true;
    const open = (raw.match(/\{/g) ?? []).length;
    const close = (raw.match(/\}/g) ?? []).length;
    if (open > close) return true;
    const salvaged = salvageVerifierSubtaskResults(raw);
    return salvaged.length > 0 && salvaged.length < subtaskCount;
  }

  private mergePreferringResolved(
    first: VerifierOutput | null,
    second: VerifierOutput,
    subtaskIds: string[],
  ): VerifierOutput {
    if (!first) return second;

    const byId = new Map<string, VerifierOutput['subtaskResults'][number]>();
    for (const id of subtaskIds) {
      const a = first.subtaskResults.find((r) => r.subtaskId === id);
      const b = second.subtaskResults.find((r) => r.subtaskId === id);
      if (b && !b.inconclusive) {
        byId.set(id, b);
      } else if (a && !a.inconclusive) {
        byId.set(id, a);
      } else {
        byId.set(id, b ?? a ?? {
          subtaskId: id,
          passed: false,
          feedback: 'Inconclusive',
          issues: [],
          inconclusive: true,
        });
      }
    }

    return normalizeVerifierOutput(
      {
        passed: false,
        feedback: second.feedback || first.feedback,
        issues: [...first.issues, ...second.issues],
        subtaskResults: subtaskIds.map((id) => byId.get(id)!),
      },
      subtaskIds,
      { fillMissingAsInconclusive: true },
    );
  }
}
