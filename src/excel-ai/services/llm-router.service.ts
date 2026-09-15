// cellix_backend/src/excel-ai/services/llm-router.service.ts

import { Injectable, Logger } from '@nestjs/common';
import { parseAgentJson } from '../../agents/utils/parse-agent-json.util';
import { AppConfigService } from '../../config/app-config.service';
import {
  buildChitchatClassifierUserMessage,
  CHITCHAT_CLASSIFIER_SYSTEM_PROMPT,
} from '../prompts/chitchat-prompt';
import { ROUTER_SYSTEM_PROMPT, buildRouterUserMessage } from '../prompts/router-system-prompt';
import { RouterDecision, RouterInput } from '../types/router.types';
import { classifyComplexity } from '../utils/complexity-classifier.util';
import { resolveLocalFindRoute } from '../utils/find-query-parser.util';
import { hasWriteIntent, isWorkbookScaffoldIntent } from '../utils/write-intent-guard.util';
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

/** Verbs that can only be a request to change the workbook. */
const WRITE_VERB =
  /\b(add|insert|create|build|generate|delete|remove|drop|clear|wipe|highlight|bold|italic|underline|sort|fill|apply|rename|merge|unmerge|split|freeze|unfreeze|hide|unhide|write|convert|replace|protect|unlock|resize|autofit|wrap|align|define|validate|dedupe|deduplicate|trim)\b/i;

/**
 * Words that are verbs in an instruction and plain nouns in a question —
 * "duplicate VALUES in column A", "what FILTER is active", "the COPY sheet".
 * Only an imperative position makes them a write. Without this split,
 * "Are there duplicate values in column A?" read as a write request.
 */
const AMBIGUOUS_WRITE_VERB =
  /(?:^|\b(?:you|please|to|and|then|also|now)\s+)(duplicate|copy|filter|format|colou?r|mark|flag|name|move|set|make|change|update|show|clean|lock|round)\b/i;

/**
 * A question about the data ("are there duplicates?", "how many blanks?",
 * "which supplier is highest?") with no verb asking for a change. The guide
 * treats these as read-only (Q&A.4) — answer them, never write. TASKS.md #214.
 */
export function isReadOnlyQuestion(message: string): boolean {
  const text = String(message ?? '').trim();
  if (!text) return false;

  const interrogative =
    /^(are|is|do|does|did|can|could|how|what|which|who|whom|whose|where|when|why|any)\b/i.test(text) ||
    /\?\s*$/.test(text);
  if (!interrogative) return false;

  // "Can you highlight the duplicates?" is a question in form and a write in
  // substance — the verb decides, not the question mark.
  return !WRITE_VERB.test(text) && !AMBIGUOUS_WRITE_VERB.test(text);
}

@Injectable()
export class LlmRouterService {
  private readonly logger = new Logger(LlmRouterService.name);

  constructor(
    private readonly openRouter: OpenRouterService,
    private readonly config: AppConfigService,
  ) {}

  /**
   * Pre-tier-classification gate (called BEFORE route()): a single LOW-tier
   * call that separates chitchat (greetings/small talk/identity questions)
   * from anything referencing the spreadsheet. CHITCHAT skips workbook
   * context load, TOON compression, and the Tier 0-3 dispatch entirely —
   * routed instead to ChitchatService.
   *
   * Fail-open to TASK on any classifier failure (throw or malformed reply):
   * worst case is one wasted tiering pass on a real greeting, which is
   * cheaper than a real task getting stuck in chitchat mode.
   */
  async classifyIntent(message: string): Promise<'CHITCHAT' | 'TASK'> {
    try {
      const raw = await this.openRouter.complete({
        systemPrompt: CHITCHAT_CLASSIFIER_SYSTEM_PROMPT,
        userMessage: buildChitchatClassifierUserMessage(message),
        model: this.config.openRouterModelLow,
        tier: 'low',
        temperature: 0,
        maxTokens: 8,
        reasoningEffort: 'none',
      });

      const label = raw.trim().toUpperCase();
      if (label.startsWith('CHITCHAT')) return 'CHITCHAT';
      if (label.startsWith('TASK')) return 'TASK';

      this.logger.warn(`classifyIntent: unrecognized label "${raw.trim()}" — defaulting to TASK`);
      return 'TASK';
    } catch (err) {
      this.logger.warn('classifyIntent call failed — defaulting to TASK', err as Error);
      return 'TASK';
    }
  }

  /**
   * Route a user message to the correct handler path.
   *
   * Priority:
   * 1. Regex fast lane (0ms) — unambiguous layout commands
   * 2. Find + copy/export — FindExportService (must beat read-only data lane)
   * 3. Complexity / workbook scaffold (write) — before data, so "total amount" cols can't trap builds
   * 4. Data query fast lane — LLM read-only path with column slicing
   * 5. Ask/plan mode short-circuit — non-data messages
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

    // Write patterns before data: "dashboard" + "total amount" (column name) would otherwise
    // false-positive into SmartDataQuery and answer "no sheet data".
    // A question with no write verb is a data query however it matches below:
    // "Are there duplicate values in column A?" hit the DUPLICATE_CHECK
    // complexity regex here, short-circuiting to write with confidence 1.0
    // before ensureWriteComplexity could downgrade it, and Tier 2 answered the
    // yes/no question by painting a conditional-format rule onto the sheet.
    // TASKS.md #214.
    if (input.mode === 'action' && !isReadOnlyQuestion(input.message)) {
      const complexityEarly = classifyComplexity(input.message);
      if (complexityEarly.match) {
        const { tier, actionHint } = complexityEarly.match;
        this.logger.debug(`Complexity regex match (pre-data): tier=${tier} actionHint=${actionHint}`);
        return {
          route: 'write',
          complexity: tier,
          actionHint,
          matchedBy: 'regex',
          confidence: 1.0,
          reasoning: `Complexity regex: tier=${tier} hint=${actionHint}`,
        };
      }
      if (isWorkbookScaffoldIntent(input.message)) {
        return {
          route: 'write',
          complexity: 3,
          actionHint: 'WORKBOOK_SCAFFOLD',
          matchedBy: 'regex',
          confidence: 0.95,
          reasoning: 'Workbook scaffold (multi-sheet / main / payments ledger) — Tier 3',
        };
      }
      // Sort / add / etc. with column names like "Total Amount" must not enter data lane.
      // An unrecognized write is deliberately escalated to Tier 3: the full
      // planner/executor/verifier pipeline is the safe default when we cannot
      // classify the request, so correctness wins over the cheaper lanes.
      if (hasWriteIntent(input.message)) {
        return {
          route: 'write',
          complexity: 3,
          // Decided by regex — no LLM was consulted on this path.
          matchedBy: 'regex',
          confidence: 0.9,
          reasoning: 'Write-intent heuristic (pre-data/LLM) — unclassified write escalated to Tier 3',
          overridden: true,
        };
      }
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

    // Late complexity already tried above for action mode; remaining action uses LLM.
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
      decision.route === 'shortcut'
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
    // Bare "total" matches column names like "total amount" in scaffold prompts; require query form.
    return /\b(find|search|sum|count|average|avg|max|min|duplicate|blank|lookup|how many|what is the total|grand total|the total of)\b/i.test(
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

    // The guide's Q&A.4 data questions are read-only by definition, but the
    // router sent "Are there duplicate values in column A?" to write, where
    // Tier 2 answered a yes/no question by inserting a Duplicate? column into
    // the user's sheet. A question with no write verb anywhere in it is a
    // query, whatever the router said. TASKS.md #214.
    if (isReadOnlyQuestion(message)) {
      this.logger.log(
        `Router said write for a question with no write verb — downgrading to data: "${message.slice(0, 100)}"`,
      );
      return { ...decision, route: 'data', complexity: undefined, actionHint: undefined };
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
        // Explicit model (not the bare `tier: 'low'` default) — decoupled from
        // the shared LOW tier so a router-specific model eval doesn't also
        // move multi-sheet.service.ts's summary or ambiguity clarification.
        // Defaults to openRouterModelLow: unset OPENROUTER_MODEL_ROUTER is a
        // no-op, identical behavior to before this override existed.
        model: this.config.openRouterModelRouter,
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
