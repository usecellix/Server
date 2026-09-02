import { ComplexityTier } from '../../excel-ai/utils/complexity-classifier.util';

/** Leave headroom for JSON plan output after reasoning. */
export const PLANNER_REASONING_MAX_TOKENS = 1024;

/** Last-resort ceiling when normal/retry budgets still yield empty or unparseable output. */
export const PLANNER_LAST_RESORT_MAX_TOKENS = 8192;

/** Tier-3 floor: the smallest budget a single-object Tier-3 build (e.g. one dashboard,
 * no compound clauses) is given. Below the old flat 8192 default, but still well above
 * Tier 2's 3000 — a Tier-3 request always implies multi-object planning overhead even
 * at its simplest. */
const TIER3_BASE_MAX_TOKENS = 4096;

/** Each additional detected object in a compound Tier-3 request adds this many
 * tokens of headroom, up to PLANNER_LAST_RESORT_MAX_TOKENS. Sized off TASKS.md #170's
 * measurement that a 20-subtask ledger plan needed ~2,814 content tokens against a
 * 3,072-token content floor — i.e. roughly 140 tokens of content per subtask/object,
 * plus reasoning headroom. */
const TIER3_PER_OBJECT_TOKENS = 1536;

/**
 * Nouns that name a distinct buildable "object" a Tier-3 plan has to produce a
 * subtask for — a chart is planning/verification work independent of a summary,
 * which is independent of a pivot table, etc. Deliberately narrow (mirrors
 * `SINGLE_ACTION_PATTERNS`'s tier-2/3 action hints in
 * `complexity-classifier.util.ts`) rather than a general clause splitter: the spec
 * 16 repro ("create a chart ,and analysis for purchase register a summary for
 * purchase register") has only one loose "and" but names three separate objects
 * (chart, analysis, summary), which is exactly the shape that exhausted the flat
 * budget — clause-counting alone would have scored it as 1.
 *
 * Widened after a follow-up review found "build me something showing trends by
 * region and by month, plus totals" scored 1 (no keyword hit at all) despite
 * clearly asking for a multi-cut breakdown — a genuinely compound request that
 * fell through to the 4096 base budget and the same wasted retry-then-8192
 * round trip spec 16 exists to avoid. Added the analysis-shape vocabulary that
 * request actually used: trend(s), total(s), region, breakdown, comparison, and
 * the month/quarter/year grouping words a "by X" cut implies.
 */
const TIER3_OBJECT_KEYWORDS =
  /\b(chart|graph|dashboard|pivot\s*table|summary|analysis|report|table|formula|validation|dropdown|kpi|ledger|sheet|tab|worksheet|trends?|totals?|region|breakdown|comparison|months?|quarters?|years?)\b/gi;

/**
 * "dashboard" is ambiguous in a way none of the other keywords are: it doubles
 * as the name of a target sheet, so "create a chart on Dashboard" was scoring
 * 2 (chart + dashboard) for a request that names exactly ONE buildable object
 * — dashboard here is WHERE the chart goes, not a second thing to build.
 * Matches a preposition (on/to/in/into/onto/for) immediately governing
 * "dashboard" (optionally "the ... dashboard", optionally "... sheet"), where
 * that prepositional phrase then ENDS the clause (a comma, a connector, or the
 * end of the string) rather than continuing into a verb. That last condition
 * is what tells "chart on Dashboard" and "summary on the Dashboard sheet"
 * (pure location, nothing follows) apart from "In dashboard create a chart"
 * (the spec 16 repro) — there "dashboard" is followed directly by "create",
 * not a clause boundary, which reads as introducing the dashboard-building
 * work itself rather than merely pointing at an existing sheet. Deliberately
 * does NOT touch prompts where "dashboard" isn't governed by a preposition at
 * all ("build a dashboard", "a purchase register dashboard") — those already
 * read as a genuine build request and are left counting normally.
 */
const AMBIGUOUS_DASHBOARD_TARGET_SHEET_REF =
  /\b(?:on|to|in|into|onto|for)\s+(?:the\s+)?dashboard\b(?:\s+sheet\b)?(?=\s*(?:,|$|\band\b|\bthen\b|\balso\b))/i;

/**
 * When AMBIGUOUS_DASHBOARD_TARGET_SHEET_REF matches an optional trailing
 * "sheet" ("on the Dashboard sheet"), that "sheet" is naming the SAME target,
 * not a second one — but "sheet" is independently one of TIER3_OBJECT_KEYWORDS
 * (for "add a sheet per month"-style compound signals). Left alone, "put a
 * summary on the Dashboard sheet" would still score 2 ({summary, sheet}) even
 * after "dashboard" is dropped, for exactly the reason this whole fix exists —
 * a location reference inflating the count. So every keyword consumed WITHIN
 * the matched ambiguous phrase (not just the literal word "dashboard") is
 * dropped from the distinct set, not only that one word.
 */

/**
 * Cheap proxy for "how compound is this Tier-3 request" — the number of distinct
 * buildable objects it names (see `TIER3_OBJECT_KEYWORDS`), deduplicated by
 * keyword so "sheet ... sheet ... sheet" doesn't inflate the score, and with
 * "dashboard" dropped when it only names a target sheet (see
 * `AMBIGUOUS_DASHBOARD_TARGET_SHEET_REF`). A single "build a dashboard" scores
 * 1; "create a chart on Dashboard" also scores 1 (dashboard is the target, not
 * a second object); the spec 16 repro (chart + analysis + summary, "In
 * dashboard create...") still scores 4 — its "dashboard" is followed directly
 * by "create", not a clause boundary, so it isn't treated as ambiguous.
 */
export function resolveTier3ComplexityScore(prompt: string): number {
  const matches = prompt.match(TIER3_OBJECT_KEYWORDS);
  if (!matches || matches.length === 0) return 1;
  const distinct = new Set(matches.map((m) => m.toLowerCase().replace(/\s+/g, ' ')));

  if (distinct.has('dashboard')) {
    const ambiguousRef = AMBIGUOUS_DASHBOARD_TARGET_SHEET_REF.exec(prompt);
    if (ambiguousRef) {
      const consumedKeywords = ambiguousRef[0].match(TIER3_OBJECT_KEYWORDS) ?? [];
      for (const keyword of consumedKeywords) {
        distinct.delete(keyword.toLowerCase().replace(/\s+/g, ' '));
      }
    }
  }

  return Math.max(1, distinct.size);
}

/**
 * Tier-aware completion budget for Planner calls, scaled within Tier 3 by a
 * complexity score (distinct object count, see `resolveTier3ComplexityScore`)
 * rather than a flat ceiling for every Tier-3 request.
 *
 * Tier 3 was flat 4096 with a 1024 reasoning cap, leaving ~3072 for content. A real
 * 20-subtask ledger plan measured 2,814 tokens — inside the budget by 8%, i.e.
 * effectively at the ceiling. TASKS.md #169's richer planner guidance (Lists
 * sheet, sheet positions, layout policy) pushed it over, so the FIRST attempt
 * truncated and the last-resort retry at 8192 ran the whole plan again:
 * 213s for one planner call in a live run, of which roughly half was the
 * wasted first attempt. TASKS.md #170 fixed this by budgeting flat 8192 for every
 * Tier-3 request — safe (max_tokens is a ceiling, not an allocation) but wasteful
 * to reason about, since a single "build a dashboard" and a multi-object compound
 * build were given the same headroom. Spec 16 fix #1: scale the ceiling by the
 * detected object count instead, so a simple Tier-3 request still gets a smaller
 * budget and a genuinely compound one still gets the full last-resort ceiling.
 */
export function resolvePlannerMaxTokens(
  complexity?: ComplexityTier | number,
  prompt = '',
): number {
  const tier = typeof complexity === 'number' ? complexity : 3;
  if (tier >= 3) {
    const score = resolveTier3ComplexityScore(prompt);
    const scaled = TIER3_BASE_MAX_TOKENS + (score - 1) * TIER3_PER_OBJECT_TOKENS;
    return Math.min(scaled, PLANNER_LAST_RESORT_MAX_TOKENS);
  }
  if (tier === 2) return 3000;
  return 2000;
}
