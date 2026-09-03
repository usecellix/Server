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
  // Spec 24: header-row cosmetics before generic highlight → FORMAT_RANGE, not FORMAT_MATCHING_ROWS
  {
    pattern:
      /\b(header|headers)\b.*\b(row|bg|background|fill|highlight|bold|colou?r|green|red|yellow|blue)\b|\b(highlight|fill|bg|background|bold|colou?r)\b.*\b(header|headers)\b/i,
    tier: 1,
    actionHint: 'HEADER_FORMAT',
  },
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

/**
 * Phrasings that imply more than one requested feature.
 *
 * Deliberately kept as a modest list rather than grown indefinitely: a word
 * list is never complete, and TASKS.md #158 is this codebase's own record of
 * what happens when one is leaned on (consolidation silently required a column
 * literally called "month"). The structural safety net is
 * `findHighestTierMatch` below plus the mid-flight escalation in
 * `conversation.service.ts` — those catch what these words miss. TASKS.md #165.
 */
const COMPOUND_SIGNALS =
  new RegExp(
    [
      '\\band then\\b',
      '\\bafter that\\b',
      ',\\s*(then|and)\\s',
      '\\bfor each\\b',
      '\\bfor every\\b',
      '\\bacross (all|every|each)\\b',
      // "one sheet per month", "a tab for each region" — the single most common
      // shape of a large build, and previously unmatched by anything here.
      '\\b(one|a|separate|individual)\\s+(sheet|tab|worksheet)s?\\s+(per|for)\\b',
      '\\b(sheets?|tabs?|worksheets?)\\s+for\\s+(all|each|every)\\b',
      '\\bmultiple\\s+(sheets?|tabs?|worksheets?)\\b',
      // A summary/roll-up sheet is by definition a second object.
      '\\b(summary|main|master|overview|consolidat\\w*)\\s+(sheet|tab|page)\\b',
      '\\bas well as\\b',
      '\\balong with\\b',
      '\\bplus\\s+(a|an|the)\\b',
    ].join('|'),
    'i',
  );

/** True when the message has multi-clause/compound phrasing implying more than one requested feature. */
export function hasCompoundSignals(message: string): boolean {
  return COMPOUND_SIGNALS.test(message);
}

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

/**
 * First match wins for the HINT; a tier-3 match anywhere wins the TIER.
 *
 * The bug this fixes: SINGLE_ACTION_PATTERNS is scanned in order and the lane-1
 * "highlight" pattern sits above the lane-3 "dashboard" pattern, so
 * *"build me a dashboard and highlight the overdue payments"* matched
 * `highlight` and routed a whole dashboard build into the single-action lane —
 * a thin answer, no planning, no verification. It escaped only when the
 * phrasing happened to carry a COMPOUND_SIGNAL, which that sentence does not.
 *
 * The obvious fix — take the maximum tier of every match — is WRONG, and was
 * tried first. The list's order encodes specificity, not just tier: the lane-2
 * FORMULA_GEN pattern matches the bare word "formula", so a plain
 * *"fill down the formula in column D"* (correctly COPY_FILL, lane 1) got
 * dragged to lane 2 by a keyword that describes none of its work. Loose
 * lower-tier keywords must not be able to outvote a precise earlier match.
 *
 * Tier 3 is different in kind, and that is why it alone overrides. A lane-3
 * pattern denotes a multi-object BUILD — a thing the single-action lanes have
 * no planner to construct, so being wrong there costs the user's result rather
 * than a few seconds. Lanes 1 and 2 keep their existing precedence untouched.
 *
 * The hint always comes from the first (most specific) match, so
 * `Tier0DirectService`'s capture-group re-run still resolves the same pattern.
 * TASKS.md #165.
 */
function findPatternMatch(
  message: string,
): { tier: ComplexityTier; actionHint: string } | null {
  let first: { tier: ComplexityTier; actionHint: string } | null = null;
  let sawBuildSignal = false;

  for (const { pattern, tier, actionHint } of SINGLE_ACTION_PATTERNS) {
    if (!pattern.test(message)) continue;
    if (!first) first = { tier, actionHint };
    if (tier === 3) sawBuildSignal = true;
  }

  if (!first) return null;
  return sawBuildSignal ? { tier: 3, actionHint: first.actionHint } : first;
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
  const singleActionMatch = findPatternMatch(message);

  if (COMPOUND_SIGNALS.test(message)) {
    // No single-action pattern matched, so there is nothing to escalate FROM.
    // Returning null hands the decision to the LLM router, which reads a vague
    // sentence far better than any regex can and already defaults to 3 when
    // unsure. Hard-coding 3 here was tried and reverted: it replaced a good
    // deferral with a blunt constant AND destroyed the `matchedBy` signal that
    // tells telemetry which component actually made the call. TASKS.md #165.
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
