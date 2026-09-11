import { Injectable } from '@nestjs/common';
import { AppConfigService } from '../../config/app-config.service';
import {
  ConversationTurn,
  LLMTier,
  MODEL_CONFIGS,
  ModelConfig,
  WorkbookContext,
} from '../../types/cellix.types';

const WEIGHTS = {
  cellsAffected: 2,
  formulaKeywords: 3,
  sheetCount: 5,
  crossSheetOp: 15,
  clarificationRound: -8,
  historyLength: 1,
  complexAggregation: 10,
  deleteOperation: 8,
  emptySheet: 5,
  longPrompt: 3,
};

const FORMULA_KEYWORDS = [
  'if(',
  'vlookup',
  'xlookup',
  'sumif',
  'countif',
  'pivot',
  'index(',
  'match(',
  'arrayformula',
];
const COMPLEX_AGG = [
  'pivot',
  'aggregate',
  'consolidate',
  'summarize',
  'group by',
  'histogram',
  'regression',
];
const DELETE_KEYWORDS = ['delete', 'remove', 'clear', 'erase', 'drop'];

export interface ComplexityScore {
  total: number;
  tier: LLMTier;
  breakdown: Record<string, number>;
  rationale: string;
}

export interface RoutingDecision {
  tier: LLMTier;
  model: string;
  config: ModelConfig;
  complexityScore: ComplexityScore;
  estimatedCostUsd: number;
  fallbackUsed: boolean;
}

/**
 * Shared per-call cost ceiling. Exported so other LLM call sites (e.g.
 * `PlannerAgent`, spec 16 fix #4) can enforce the same cap this router already
 * applies to its own high-tier routing decision, rather than each call site
 * inventing its own number.
 */
export const COST_CAP_USD = 0.15;

const RATE_LIMIT_RETRY_TIER: Record<LLMTier, LLMTier | null> = {
  high: 'medium',
  medium: 'low',
  low: null,
};

function buildRationale(breakdown: Record<string, number>, tier: LLMTier): string {
  const top = Object.entries(breakdown)
    .filter(([, value]) => value > 0)
    .sort(([, a], [, b]) => b - a)
    .slice(0, 2)
    .map(([key]) => key)
    .join(', ');

  return `${tier} tier — driven by: ${top || 'simple task'}`;
}

export function scoreTaskComplexity(
  prompt: string,
  context: WorkbookContext,
  conversationHistory: ConversationTurn[] = [],
): ComplexityScore {
  const lower = prompt.toLowerCase();
  const words = lower.split(/\s+/);
  const breakdown: Record<string, number> = {};

  const totalCells = context.sheets.reduce((sum, sheet) => sum + sheet.rowCount * sheet.colCount, 0);
  breakdown.cells = Math.min(30, Math.floor(totalCells / 10) * WEIGHTS.cellsAffected);

  const formulaHits = FORMULA_KEYWORDS.filter((keyword) => lower.includes(keyword));
  breakdown.formulas = Math.min(20, formulaHits.length * WEIGHTS.formulaKeywords);

  const extraSheets = Math.max(0, context.sheets.length - 1);
  breakdown.sheets = extraSheets * WEIGHTS.sheetCount;

  const crossSheet =
    context.sheets.length > 1 &&
    context.sheets.some((sheet) => lower.includes(sheet.sheetName.toLowerCase()));
  breakdown['cross-sheet'] = crossSheet ? WEIGHTS.crossSheetOp : 0;

  const clarRounds = conversationHistory.filter(
    (turn) => turn.role === 'assistant' && turn.content.startsWith('[Clarification needed]'),
  ).length;
  breakdown.clarification = clarRounds * WEIGHTS.clarificationRound;

  breakdown.history = Math.min(10, conversationHistory.length * WEIGHTS.historyLength);

  const hasComplexAgg = COMPLEX_AGG.some((keyword) => lower.includes(keyword));
  breakdown.aggregation = hasComplexAgg ? WEIGHTS.complexAggregation : 0;

  const hasDelete = DELETE_KEYWORDS.some((keyword) => words.includes(keyword));
  breakdown.delete = hasDelete ? WEIGHTS.deleteOperation : 0;

  const isEmptySheet = context.sheets.some((sheet) => sheet.rowCount <= 1 && sheet.colCount <= 1);
  breakdown['empty-sheet'] = isEmptySheet ? WEIGHTS.emptySheet : 0;

  const extraWords = Math.max(0, words.length - 15);
  breakdown['prompt-length'] = Math.min(10, Math.floor(extraWords / 20) * WEIGHTS.longPrompt);

  const raw = Object.values(breakdown).reduce((sum, value) => sum + value, 0);
  let total = Math.max(0, Math.min(100, raw));

  // Creating data on an empty sheet is never a low-tier task — needs reliable JSON actions.
  const isPopulate =
    isEmptySheet &&
    /\b(create|generate|populate|dummy|sample|fill|add|insert|build|make)\b/.test(lower);
  if (isPopulate && total < 31) {
    total = 31;
    breakdown['populate-boost'] = 31 - raw;
  }

  const tier: LLMTier = total <= 30 ? 'low' : total <= 65 ? 'medium' : 'high';

  return {
    total,
    tier,
    breakdown,
    rationale: buildRationale(breakdown, tier),
  };
}

/**
 * Cost estimate for one LLM call at a given prompt/completion token count.
 * Exported (renamed from the router-only `estimateCostUsd`) so a caller with
 * its own completion-token budget — e.g. a Planner call whose `maxTokens` is
 * NOT the tier's flat `ModelConfig.maxTokens` — can price the call it is
 * actually about to make, rather than the router's own fixed-budget estimate.
 */
export function estimateLlmCallCostUsd(
  pricing: Pick<ModelConfig, 'costPer1kPrompt' | 'costPer1kCompletion'>,
  promptTokens: number,
  completionTokens: number,
): number {
  const estimatedPromptCost = (promptTokens / 1000) * pricing.costPer1kPrompt;
  const estimatedCompletionCost = (completionTokens / 1000) * pricing.costPer1kCompletion;
  return estimatedPromptCost + estimatedCompletionCost;
}

function estimateCostUsd(config: ModelConfig, promptTokenEstimate: number): number {
  return estimateLlmCallCostUsd(config, promptTokenEstimate, config.maxTokens);
}

export interface ResolvedModelPricing {
  pricing: ModelConfig;
  tier: LLMTier;
  /** True when `model` did not match any of the LOW/MEDIUM/HIGH configured
   * models and HIGH pricing was used as a conservative fallback — the caller
   * should log this so the estimate is understood as approximate, not exact. */
  approximate: boolean;
}

/**
 * Resolve pricing for an arbitrary model string by matching it against the
 * models actually configured for the three LLMTiers (`AppConfigService`), NOT
 * by assuming a call is always HIGH tier. Needed because call sites like
 * `PlannerAgent` can be pointed at a different model than
 * `openRouterModelHigh` via a per-purpose override (`openRouterModelPlanner`,
 * spec 16 fix #2) — pricing that model as HIGH unconditionally would silently
 * misprice it the moment that override diverges from HIGH's actual model.
 *
 * A model matching none of the three configured models (e.g. an eval pointed
 * at a model with no pricing entry here) falls back to HIGH pricing — the most
 * expensive of the three — as a conservative default so an unrecognized model
 * can never be UNDER-estimated into slipping past the cost cap. `approximate:
 * true` flags this so the caller can log it rather than presenting the number
 * as exact.
 */
export function resolvePricingForModel(
  model: string,
  config: Pick<AppConfigService, 'openRouterModelLow' | 'openRouterModelMedium' | 'openRouterModelHigh'>,
): ResolvedModelPricing {
  if (model === config.openRouterModelLow) {
    return { pricing: MODEL_CONFIGS.low, tier: 'low', approximate: false };
  }
  if (model === config.openRouterModelMedium) {
    return { pricing: MODEL_CONFIGS.medium, tier: 'medium', approximate: false };
  }
  if (model === config.openRouterModelHigh) {
    return { pricing: MODEL_CONFIGS.high, tier: 'high', approximate: false };
  }
  return { pricing: MODEL_CONFIGS.high, tier: 'high', approximate: true };
}

@Injectable()
export class ModelRouter {
  private readonly rateLimitedTiers = new Set<LLMTier>();

  constructor(private readonly config: AppConfigService) {}

  route(
    prompt: string,
    context: WorkbookContext,
    conversationHistory: ConversationTurn[],
    promptTokenEstimate: number,
  ): RoutingDecision {
    const complexity = scoreTaskComplexity(prompt, context, conversationHistory);
    let tier: LLMTier = complexity.tier;
    let fallbackUsed = false;

    while (this.rateLimitedTiers.has(tier)) {
      const fallback = RATE_LIMIT_RETRY_TIER[tier];
      if (!fallback) {
        throw new Error('All model tiers are rate-limited. Please try again shortly.');
      }
      tier = fallback;
      fallbackUsed = true;
    }

    let activeConfig = MODEL_CONFIGS[tier];
    let estimatedCostUsd = estimateCostUsd(activeConfig, promptTokenEstimate);

    if (estimatedCostUsd > COST_CAP_USD && tier === 'high') {
      tier = 'medium';
      fallbackUsed = true;
      activeConfig = MODEL_CONFIGS[tier];
      estimatedCostUsd = estimateCostUsd(activeConfig, promptTokenEstimate);
    }

    const model = this.resolveModelId(tier);

    return {
      tier,
      model,
      config: { ...activeConfig, model },
      complexityScore: complexity,
      estimatedCostUsd,
      fallbackUsed,
    };
  }

  markRateLimited(tier: LLMTier): void {
    this.rateLimitedTiers.add(tier);
    setTimeout(() => {
      this.rateLimitedTiers.delete(tier);
    }, 60_000);
  }

  resolveModelId(tier: LLMTier): string {
    const model =
      tier === 'low'
        ? this.config.openRouterModelLow
        : tier === 'high'
          ? this.config.openRouterModelHigh
          : this.config.openRouterModelMedium;

    if (model.includes('openrouter/auto')) {
      return tier === 'high' ? 'openai/gpt-5' : 'openai/gpt-5-mini';
    }

    return model;
  }

  buildThinkingMessage(routing: RoutingDecision): string {
    const { complexityScore, tier, fallbackUsed } = routing;
    if (fallbackUsed) {
      return `Using ${tier} model (fallback active, complexity: ${complexityScore.total}/100)…`;
    }
    return `Routing to ${tier} model (complexity: ${complexityScore.total}/100)…`;
  }
}
