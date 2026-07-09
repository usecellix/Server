// cellix_backend/src/excel-ai/services/llm-router.service.ts

import { Injectable, Logger } from '@nestjs/common';
import { parseAgentJson } from '../../agents/utils/parse-agent-json.util';
import { AppConfigService } from '../../config/app-config.service';
import { ROUTER_SYSTEM_PROMPT, buildRouterUserMessage } from '../prompts/router-system-prompt';
import { RouterDecision, RouterInput } from '../types/router.types';
import { OpenRouterService } from './openrouter.service';

// Regex fast lane — these NEVER go to the LLM router.
// Zero ambiguity + regex is 0ms vs 100ms LLM.
const INSTANT_SHORTCUT_PATTERNS: Array<{ pattern: RegExp; action: string }> = [
  { pattern: /freeze\s+(top\s+row|first\s+row|row\s+1)/i, action: 'FREEZE_PANES' },
  { pattern: /freeze\s+(first\s+col|left\s+col|column\s+a)/i, action: 'FREEZE_PANES' },
  { pattern: /unfreeze/i, action: 'UNFREEZE_PANES' },
  { pattern: /protect\s+(this\s+)?sheet/i, action: 'PROTECT_SHEET' },
  { pattern: /unprotect\s+(this\s+)?sheet/i, action: 'UNPROTECT_SHEET' },
  { pattern: /zoom\s+to\s+\d+%?/i, action: 'SET_ZOOM' },
  { pattern: /set\s+zoom\s+to\s+\d+%?/i, action: 'SET_ZOOM' },
];

@Injectable()
export class LlmRouterService {
  private readonly logger = new Logger(LlmRouterService.name);

  constructor(
    private readonly openRouter: OpenRouterService,
    private readonly config: AppConfigService,
  ) {}

  /**
   * Route a user message to the correct handler path.
   *
   * Priority:
   * 1. Regex fast lane (0ms) — unambiguous layout commands
   * 2. Data query fast lane — LLM read-only path with column slicing
   * 3. Ask/plan mode short-circuit — non-data messages
   * 4. LLM Router (LOW tier, ~100ms) — everything else
   */
  async route(input: RouterInput): Promise<RouterDecision> {
    const instantMatch = this.tryInstantShortcut(input.message);
    if (instantMatch) {
      this.logger.debug(`Instant shortcut match: ${instantMatch.action}`);
      return {
        route: 'shortcut',
        action: instantMatch.action,
        confidence: 1.0,
        reasoning: 'Matched instant shortcut regex — no LLM needed',
      };
    }

    if (this.quickDataCheck(input.message)) {
      return {
        route: 'data',
        confidence: 0.85,
        reasoning: 'Matched data query keywords — SmartDataQuery (MEDIUM tier)',
      };
    }

    if (input.mode !== 'action') {
      return {
        route: 'ask',
        confidence: 0.9,
        reasoning: 'Non-action mode — routing to ask path',
      };
    }

    return this.callLlmRouter(input);
  }

  private tryInstantShortcut(message: string): { action: string } | null {
    for (const { pattern, action } of INSTANT_SHORTCUT_PATTERNS) {
      if (pattern.test(message)) return { action };
    }
    return null;
  }

  private quickDataCheck(message: string): boolean {
    return /\b(find|search|sum|count|average|avg|max|min|duplicate|blank|lookup|how many|total)\b/i.test(
      message,
    );
  }

  private async callLlmRouter(input: RouterInput): Promise<RouterDecision> {
    const userMessage = buildRouterUserMessage(
      input.message,
      input.activeSheet,
      input.sheetHeaders,
      input.recentHistory,
      input.mode,
    );

    try {
      const raw = await this.openRouter.complete({
        systemPrompt: ROUTER_SYSTEM_PROMPT,
        userMessage,
        tier: 'low',
        temperature: 0,
        maxTokens: 256,
        reasoningEffort: 'none',
      });

      const parsed = parseAgentJson<RouterDecision>(raw);

      if (!parsed?.route) {
        this.logger.warn('LLM Router returned invalid JSON — defaulting to write');
        return this.fallbackDecision(input.message);
      }

      this.logger.debug(
        `Router decision: route=${parsed.route} confidence=${parsed.confidence} reason="${parsed.reasoning}"`,
      );

      return parsed;
    } catch (err) {
      this.logger.error('LLM Router call failed', err);
      return this.fallbackDecision(input.message);
    }
  }

  /**
   * If the LLM Router fails entirely, fall back to a safe default.
   * Prefer 'ask' over 'write' to avoid unintended modifications.
   */
  private fallbackDecision(message: string): RouterDecision {
    const looksLikeWrite = /\b(create|add|delete|remove|sort|format|bold|color|fill|write|insert|rename|copy)\b/i.test(
      message,
    );
    return {
      route: looksLikeWrite ? 'write' : 'ask',
      confidence: 0.4,
      reasoning: 'LLM Router fallback — regex heuristic',
    };
  }
}
