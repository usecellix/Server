// cellix_backend/src/excel-ai/utils/complexity-classifier.util.ts

export type ComplexityTier = 0 | 1 | 2 | 3;

export interface ComplexityMatch {
  tier: ComplexityTier;
  actionHint: string;
  matchedBy: 'regex' | 'llm-fallback';
  confidence?: number;
}

export interface ComplexityClassifierResult {
  match: ComplexityMatch | null;
}

const SINGLE_ACTION_PATTERNS: Array<{ pattern: RegExp; tier: ComplexityTier; actionHint: string }> = [
  // Tier 0 — explicit target, pure structural/cosmetic, zero interpretation
  { pattern: /\b(bold|italic|underline)\b.*\b[a-z]+\d+(:[a-z]+\d+)?\b/i, tier: 0, actionHint: 'CELL_FORMAT' },
  { pattern: /\bfreeze\s+(top\s+)?row\b/i, tier: 0, actionHint: 'FREEZE_PANES' },
  { pattern: /\b(hide|unhide|show)\s+(column|row|sheet)\b/i, tier: 0, actionHint: 'VISIBILITY_TOGGLE' },
  { pattern: /\b(insert|delete)\s+(a\s+)?(row|column)\b/i, tier: 0, actionHint: 'ROW_COL_STRUCTURE' },

  // Tier 1 — single LLM call, no verification, low stakes
  { pattern: /\b(sort|filter)\b.*\bby\b/i, tier: 1, actionHint: 'SORT_OR_FILTER' },
  { pattern: /\bfind\s*(and)?\s*replace\b/i, tier: 1, actionHint: 'FIND_REPLACE' },
  { pattern: /\b(highlights?|conditional formats?)\b/i, tier: 1, actionHint: 'CONDITIONAL_FORMAT' },
  {
    pattern: /\b(remove|remvoe|clear|unhighlight)\b.*\b(highlights?|fills?|colou?rs?)\b/i,
    tier: 1,
    actionHint: 'CONDITIONAL_FORMAT',
  },
  {
    pattern: /\b(highlights?|fills?|colou?rs?)\b.*\b(remove|clear)\b/i,
    tier: 1,
    actionHint: 'CONDITIONAL_FORMAT',
  },
  { pattern: /\bfill\s+down\b|\bcopy\s+format(ting)?\b/i, tier: 1, actionHint: 'COPY_FILL' },

  // Tier 3 — multi-object reports require planning across formulas, formatting, and charts
  { pattern: /\bdashboard\b/i, tier: 3, actionHint: 'DASHBOARD' },

  // Tier 2 — formula/computation/structured object, verification mandatory
  { pattern: /\bcalculate\b.*%|=|\bformula\b|\bif\s.*then\b/i, tier: 2, actionHint: 'FORMULA_GEN' },
  { pattern: /\bpivot\s*table\b/i, tier: 2, actionHint: 'PIVOT_TABLE' },
  { pattern: /\bcharts?\b|\bgraphs?\b/i, tier: 2, actionHint: 'CHART' },
  { pattern: /\bduplicate\b/i, tier: 2, actionHint: 'DUPLICATE_CHECK' },
  { pattern: /\bvalidation\b|\bdropdown\b/i, tier: 2, actionHint: 'DATA_VALIDATION' },
  { pattern: /#(REF|N\/A|VALUE|DIV\/0)!?/i, tier: 2, actionHint: 'ERROR_FIX' },
];

const COMPOUND_SIGNALS =
  /\band then\b|\bafter that\b|,\s*(then|and)\s|\bfor each sheet\b|\bacross (all|every) sheets?\b/i;

export const NUMERIC_FINANCIAL_HINT = /\b(gst|gstin|amount|total|balance|invoice|tax|₹|rs\.?)\b/i;

const TIER0_ACTION_HINTS = new Set(['CELL_FORMAT', 'FREEZE_PANES', 'VISIBILITY_TOGGLE', 'ROW_COL_STRUCTURE']);

/** Re-run the tier-0 classifier pattern for capture groups used by Tier0DirectService. */
export function extractTier0PatternMatch(
  message: string,
  actionHint: string,
): RegExpMatchArray | null {
  if (!TIER0_ACTION_HINTS.has(actionHint)) {
    return null;
  }

  for (const { pattern, tier, actionHint: hint } of SINGLE_ACTION_PATTERNS) {
    if (tier !== 0 || hint !== actionHint) {
      continue;
    }
    const match = pattern.exec(message);
    if (match) {
      return match;
    }
  }

  return null;
}

function findFirstSingleActionMatch(
  message: string,
): { tier: ComplexityTier; actionHint: string } | null {
  for (const { pattern, tier, actionHint } of SINGLE_ACTION_PATTERNS) {
    if (pattern.test(message)) {
      return { tier, actionHint };
    }
  }
  return null;
}

function applyFindReplaceEscalation(
  tier: ComplexityTier,
  actionHint: string,
  message: string,
): { tier: ComplexityTier; actionHint: string } {
  if (tier === 1 && actionHint === 'FIND_REPLACE' && NUMERIC_FINANCIAL_HINT.test(message)) {
    return { tier: 2, actionHint: 'FIND_REPLACE' };
  }
  return { tier, actionHint };
}

export function classifyComplexity(
  message: string,
  _activeSheetContext?: { hasHeaders?: boolean },
): ComplexityClassifierResult {
  const singleActionMatch = findFirstSingleActionMatch(message);

  if (COMPOUND_SIGNALS.test(message)) {
    if (!singleActionMatch) {
      return { match: null };
    }

    return {
      match: {
        tier: 3,
        actionHint: singleActionMatch.actionHint,
        matchedBy: 'regex',
      },
    };
  }

  if (!singleActionMatch) {
    return { match: null };
  }

  const { tier, actionHint } = applyFindReplaceEscalation(
    singleActionMatch.tier,
    singleActionMatch.actionHint,
    message,
  );

  return {
    match: {
      tier,
      actionHint,
      matchedBy: 'regex',
    },
  };
}
