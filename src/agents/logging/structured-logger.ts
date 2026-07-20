import { Injectable, Logger } from '@nestjs/common';
import { AgentLogEvent, DomainToolLog, TierDecisionLog } from '../types/log.types';

const SLOW_CALL_MS = 15_000;

@Injectable()
export class StructuredLogger {
  private readonly logger = new Logger(StructuredLogger.name);

  logAgentEvent(event: AgentLogEvent): void {
    const sanitized = this.sanitizeEvent(event);
    this.logger.log(JSON.stringify({ event: 'agent_call', ...sanitized }));
    if (sanitized.durationMs > SLOW_CALL_MS) {
      this.logger.warn(JSON.stringify({ event: 'agent_slow_call', ...sanitized }));
    }
  }

  logTierDecision(event: TierDecisionLog): void {
    const sanitized: TierDecisionLog = {
      ...event,
      message: this.redactText(event.message),
      actionHint: this.redactText(event.actionHint),
    };
    this.logger.log(JSON.stringify({ event: 'tier_decision', ...sanitized }));
  }

  logDomainToolCall(event: DomainToolLog): void {
    this.logger.log(
      JSON.stringify({
        event: 'domain_tool_call',
        ...event,
        error: typeof event.error === 'string' ? this.redactText(event.error) : undefined,
      }),
    );
  }

  debugRawResponse(
    correlationId: string,
    agent: AgentLogEvent['agent'],
    model: string,
    rawResponse: string,
  ): void {
    this.logger.debug(
      JSON.stringify({
        event: 'agent_raw_response',
        correlationId,
        agent,
        model,
        rawResponse: this.redactText(rawResponse),
      }),
    );
  }

  warnParseFailure(
    correlationId: string,
    agent: AgentLogEvent['agent'],
    model: string,
    rawResponse: string,
    parseError: string,
  ): void {
    this.logger.warn(
      JSON.stringify({
        event: 'agent_parse_failure',
        correlationId,
        agent,
        model,
        parseError: this.redactText(parseError),
        rawResponse: this.redactText(rawResponse),
      }),
    );
  }

  estimateTokens(raw: string): number {
    return Math.max(0, Math.ceil((raw?.length ?? 0) / 4));
  }

  private sanitizeEvent(event: AgentLogEvent): AgentLogEvent {
    return {
      ...event,
      rawResponse:
        typeof event.rawResponse === 'string' ? this.redactText(event.rawResponse) : undefined,
      parsedResponse: this.redactUnknown(event.parsedResponse),
      error: typeof event.error === 'string' ? this.redactText(event.error) : undefined,
    };
  }

  private redactUnknown(value: unknown): unknown {
    if (typeof value === 'string') return this.redactText(value);
    if (Array.isArray(value)) return value.map((item) => this.redactUnknown(item));
    if (value && typeof value === 'object') {
      const next: Record<string, unknown> = {};
      for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
        if (/(api[_-]?key|token|password|authorization)/i.test(key)) {
          next[key] = '[REDACTED]';
          continue;
        }
        next[key] = this.redactUnknown(entry);
      }
      return next;
    }
    return value;
  }

  private redactText(input: string): string {
    return input
      .replace(/(Bearer\s+)[A-Za-z0-9\-._~+/]+=*/gi, '$1[REDACTED]')
      .replace(/(sk-[A-Za-z0-9_-]{8,})/g, '[REDACTED_API_KEY]')
      .replace(
        /((?:api[_-]?key|token|password|authorization)\s*[:=]\s*["']?)([^"'\s,}]+)/gi,
        '$1[REDACTED]',
      );
  }
}
