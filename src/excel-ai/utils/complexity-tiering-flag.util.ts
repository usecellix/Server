/**
 * Rollout control for complexity-tiered write dispatch (Spec 08).
 *
 * Modes:
 * - off     — always execute via Tier 3 orchestrator (production default)
 * - shadow  — classify + log, still always Tier 3 (collect TierDecisionLog data)
 * - tier01  — enable Tier 0/1 handlers only; Tier 2+ still orchestrator
 * - full    — enable Tier 0–2 handlers (Tier 3 unchanged)
 *
 * Env: ENABLE_COMPLEXITY_TIERING=off|shadow|tier01|full
 * Aliases: false/0 → off; true/on/1 → full
 */
export type ComplexityTieringMode = 'off' | 'shadow' | 'tier01' | 'full';

export function parseComplexityTieringMode(
  raw: string | undefined,
  nodeEnv = process.env.NODE_ENV,
): ComplexityTieringMode {
  const value = (raw ?? '').trim().toLowerCase();

  if (!value) {
    // Production defaults off; local/test default full so existing tier tests keep working.
    return nodeEnv === 'production' ? 'off' : 'full';
  }

  if (value === 'off' || value === 'false' || value === '0') return 'off';
  if (value === 'shadow') return 'shadow';
  if (value === 'tier01' || value === 'tier0-1' || value === 'tier0_1') return 'tier01';
  if (value === 'full' || value === 'on' || value === 'true' || value === '1') return 'full';

  return nodeEnv === 'production' ? 'off' : 'full';
}

export function getComplexityTieringMode(): ComplexityTieringMode {
  return parseComplexityTieringMode(process.env.ENABLE_COMPLEXITY_TIERING);
}

/**
 * Given the classifier's tier, return the tier path that should actually execute.
 */
export function resolveExecutableTier(
  classifiedTier: 0 | 1 | 2 | 3,
  mode: ComplexityTieringMode = getComplexityTieringMode(),
): 0 | 1 | 2 | 3 {
  if (mode === 'off' || mode === 'shadow') {
    return 3;
  }
  if (mode === 'tier01') {
    return classifiedTier <= 1 ? classifiedTier : 3;
  }
  return classifiedTier;
}

export function isShadowTieringMode(
  mode: ComplexityTieringMode = getComplexityTieringMode(),
): boolean {
  return mode === 'shadow';
}
