/**
 * P0.2 — advice guardrail (spec §1.4).
 *
 * CELLIX computes and flags; it never gives tax/legal verdicts. This module is
 * deterministic (no LLM) so it holds even when the model ignores the prompt:
 *
 *  1. `isAdviceShapedQuestion` — "can I claim…", "should I file…", "is it allowed…".
 *  2. `containsVerdictLanguage` — "you can claim", "you should file", "it is allowed"…
 *  3. `applyAdviceGuardrail` — for advice-shaped questions, or answers that contain
 *     verdict language, soften the verdict wording and append the fixed line.
 *
 * Plain computations ("total of column D") are untouched.
 */

export const ADVICE_DISCLAIMER =
  'This is a computation based on encoded rules. The decision rests with you as the professional.';

/** User asks for a decision/permission rather than a computation. */
const ADVICE_QUESTION_PATTERNS: RegExp[] = [
  /\b(can|could|may|am)\s+(i|we)\s+(claim|take|avail|set\s*off|offset|deduct|file|skip|ignore|reverse|carry\s*forward|adjust|write\s*off)\b/i,
  /\b(should|shall|must|do)\s+(i|we)\s+(claim|file|pay|deduct|reverse|reject|accept|register|charge|deposit|ignore|write\s*off|need\s+to)\b/i,
  /\bis\s+(it|this|that)\s+(allowed|permitted|legal|illegal|valid|deductible|eligible|taxable|exempt|compliant|mandatory|required|safe)\b/i,
  /\b(am|are)\s+(i|we)\s+(eligible|liable|required|allowed|entitled|exempt|covered|compliant)\b/i,
  /\bis\s+(there|any)\s+(penalty|interest|liability|risk)\b/i,
  /\b(will|would)\s+(i|we)\s+(get|face|incur|be\s+liable|be\s+penali[sz]ed)\b/i,
  /\b(is|are)\s+\w+(\s+\w+){0,4}\s+(allowed|deductible|claimable|eligible|taxable|exempt|liable)\b/i,
];

/** Phrases that turn a computed fact into a verdict. */
const VERDICT_PATTERNS: Array<[RegExp, string]> = [
  [/\byou\s+(can|may|are\s+allowed\s+to|are\s+entitled\s+to|are\s+eligible\s+to)\s+(safely\s+)?(claim|take|avail|deduct|set\s*off|offset)\b/gi, 'the computed figures indicate this may be claimable — confirm the position'],
  [/\byou\s+(should|must|need\s+to|have\s+to)\s+(file|pay|deposit|reverse|reject|accept|register|report)\b/gi, 'the computed figures flag this for filing/payment review'],
  [/\b(it|this|that)\s+(is|’s|'s)\s+(allowed|permitted|legal|compliant|deductible|eligible)\b/gi, 'the rules checked do not flag this'],
  [/\b(it|this|that)\s+(is|’s|'s)\s+(not\s+allowed|illegal|non-?compliant|not\s+deductible|ineligible)\b/gi, 'the rules checked flag this'],
  [/\bfile\s+(this|it|now)\b/gi, 'review this for filing'],
  // Instruction-style filing/claiming guidance ("file by the due date", "avoid filing if…").
  [/\b(?:you\s+)?(?:should\s+|must\s+|need\s+to\s+)?file\s+by\b/gi, 'the filing due date applies; check'],
  [/\b(?:you\s+)?(?:should\s+|must\s+)?(avoid|refrain\s+from|do\s+not|don['’]t)\s+(filing|claiming|paying|depositing)\b/gi, 'the computed flags may be relevant to $2'],
  [/\bas\s+a\s+general\s+rule,?\s+/gi, 'For reference, '],
  [/\byou('re|\s+are)\s+(fine|safe|good\s+to\s+go|in\s+the\s+clear)\b/gi, 'no flags were raised by the rules checked'],
];

export function isAdviceShapedQuestion(message: string | undefined | null): boolean {
  if (!message) return false;
  return ADVICE_QUESTION_PATTERNS.some((re) => re.test(message));
}

export function containsVerdictLanguage(text: string | undefined | null): boolean {
  if (!text) return false;
  return VERDICT_PATTERNS.some(([re]) => {
    re.lastIndex = 0;
    return re.test(text);
  });
}

export function softenVerdicts(text: string): string {
  let out = text;
  for (const [re, replacement] of VERDICT_PATTERNS) {
    out = out.replace(re, replacement);
  }
  return out;
}

export interface AdviceGuardrailResult {
  answer: string;
  /** True when the disclaimer was appended and/or verdict wording was softened. */
  guarded: boolean;
}

/**
 * Apply the guardrail to a final answer. `userMessage` is the question that
 * produced it. Idempotent: the disclaimer is never added twice.
 */
export function applyAdviceGuardrail(
  answer: string,
  userMessage: string | undefined | null,
): AdviceGuardrailResult {
  if (typeof answer !== 'string' || answer.length === 0) return { answer, guarded: false };

  const adviceQuestion = isAdviceShapedQuestion(userMessage);
  const verdict = containsVerdictLanguage(answer);
  if (!adviceQuestion && !verdict) return { answer, guarded: false };

  let out = verdict ? softenVerdicts(answer) : answer;
  if (!out.includes(ADVICE_DISCLAIMER)) {
    out = `${out.trimEnd()}\n\n_${ADVICE_DISCLAIMER}_`;
  }
  return { answer: out, guarded: true };
}
