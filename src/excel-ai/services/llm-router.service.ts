// cellix_backend/src/excel-ai/services/llm-router.service.ts

import { Injectable, Logger } from '@nestjs/common';
import { parseAgentJson } from '../../agents/utils/parse-agent-json.util';
import { AppConfigService } from '../../config/app-config.service';
import { ROUTER_SYSTEM_PROMPT, buildRouterUserMessage } from '../prompts/router-system-prompt';
import { RouterDecision, RouterInput } from '../types/router.types';
import { classifyComplexity } from '../utils/complexity-classifier.util';
import { resolveLocalFindRoute } from '../utils/find-query-parser.util';
import { hasWriteIntent } from '../utils/write-intent-guard.util';
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

function isValidComplexity(value: unknown): value is 0 | 1 | 2 | 3 {
  return value === 0 || value === 1 || value === 2 || value === 3;
}

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
   * 2. Find + copy/export — FindExportService (must beat read-only data lane)
   * 3. Data query fast lane — LLM read-only path with column slicing
   * 4. Ask/plan mode short-circuit — non-data messages
   * 5. Complexity regex lane — write-tier classification before LLM
   * 6. LLM Router (LOW tier, ~100ms) — everything else
   * 7. Write-intent guard — escalate misrouted mutations to write (never silent read-only)
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

    // "find X and copy those rows to a new sheet" must not hit the read-only data lane
    // (quickDataCheck matches on "find" alone and would skip FindExportService).
    if (resolveLocalFindRoute(input.message) === 'export_rows') {
      return this.applyWriteIntentGuard(
        {
          route: 'export',
          confidence: 0.95,
          reasoning: 'Matched find + copy/export to sheet — FindExportService',
        },
        input.message,
      );
    }

    if (this.quickDataCheck(input.message)) {
      return this.applyWriteIntentGuard(
        {
          route: 'data',
          confidence: 0.85,
          reasoning: 'Matched data query keywords — SmartDataQuery (MEDIUM tier)',
        },
        input.message,
      );
    }

    if (input.mode !== 'action') {
      return this.applyWriteIntentGuard(
        {
          route: 'ask',
          confidence: 0.9,
          reasoning: 'Non-action mode — routing to ask path',
        },
        input.message,
      );
    }

    const complexityResult = classifyComplexity(input.message);
    if (complexityResult.match) {
      const { tier, actionHint } = complexityResult.match;
      this.logger.debug(`Complexity regex match: tier=${tier} actionHint=${actionHint}`);
      return {
        route: 'write',
        complexity: tier,
        actionHint,
        matchedBy: 'regex',
        confidence: 1.0,
        reasoning: `Complexity regex: tier=${tier} hint=${actionHint}`,
      };
    }

    const llmDecision = await this.callLlmRouter(input);
    return this.applyWriteIntentGuard(llmDecision, input.message);
  }

  /**
   * Safety net: if classification picked data/ask/etc. but the message
   * deterministically implies a sheet mutation, escalate to write.
   * Shortcut and export keep their dedicated pipelines.
   */
  private applyWriteIntentGuard(decision: RouterDecision, message: string): RouterDecision {
    if (
      decision.route === 'write' ||
      decision.route === 'shortcut' ||
      decision.route === 'export'
    ) {
      return decision;
    }

    if (!hasWriteIntent(message)) {
      return decision;
    }

    this.logger.warn('write-intent-guard: overriding route', {
      original: decision.route,
      message,
    });

    const complexityResult = classifyComplexity(message);
    const complexity =
      complexityResult.match?.tier ??
      (typeof decision.complexity === 'number' ? decision.complexity : 3);

    return {
      ...decision,
      route: 'write',
      complexity,
      actionHint: complexityResult.match?.actionHint ?? decision.actionHint,
      matchedBy: complexityResult.match?.matchedBy ?? decision.matchedBy ?? 'llm-fallback',
      confidence: Math.max(decision.confidence, 0.85),
      reasoning: `write-intent-guard: overridden from ${decision.route} — ${decision.reasoning}`,
      overridden: true,
    };
  }

  /**
   * Instant regex shortcut check (no LLM). Used to skip SheetAnalyzer for layout commands.
   */
  peekInstantShortcut(message: string): { action: string } | null {
    return this.tryInstantShortcut(message);
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

  private ensureWriteComplexity(
    decision: RouterDecision,
    message: string,
    source: 'llm' | 'fallback',
  ): RouterDecision {
    if (decision.route !== 'write') {
      return decision;
    }

    const complexity = isValidComplexity(decision.complexity) ? decision.complexity : 3;
    const matchedBy = decision.matchedBy ?? 'llm-fallback';

    if (source === 'llm' && matchedBy === 'llm-fallback') {
      this.logger.warn(
        `Complexity regex miss — LLM fallback: message="${message.slice(0, 120)}" tier=${complexity} actionHint=${decision.actionHint ?? 'none'}`,
      );
    }

    return {
      ...decision,
      complexity,
      matchedBy,
    };
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
        return this.ensureWriteComplexity(this.fallbackDecision(input.message), input.message, 'fallback');
      }

      this.logger.debug(
        `Router decision: route=${parsed.route} confidence=${parsed.confidence} complexity=${parsed.complexity ?? 'none'} reason="${parsed.reasoning}"`,
      );

      return this.ensureWriteComplexity(parsed, input.message, 'llm');
    } catch (err) {
      this.logger.error('LLM Router call failed', err);
      return this.ensureWriteComplexity(this.fallbackDecision(input.message), input.message, 'fallback');
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
