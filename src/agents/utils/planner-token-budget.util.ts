import { ComplexityTier } from '../../excel-ai/utils/complexity-classifier.util';

/** Leave headroom for JSON plan output after reasoning. */
export const PLANNER_REASONING_MAX_TOKENS = 1024;

/** Last-resort ceiling when normal/retry budgets still yield empty or unparseable output. */
export const PLANNER_LAST_RESORT_MAX_TOKENS = 8192;

/**
 * Tier-aware completion budget for Planner calls.
 * Higher tiers (compound / dashboard) need room for both reasoning and JSON output.
 */
export function resolvePlannerMaxTokens(complexity?: ComplexityTier | number): number {
  const tier = typeof complexity === 'number' ? complexity : 3;
  if (tier >= 3) return 4096;
  if (tier === 2) return 3000;
  return 2000;
}
