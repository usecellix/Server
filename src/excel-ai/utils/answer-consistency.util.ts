/**
 * Internal-consistency checking for read-only (Ask mode) answers.
 *
 * Task #90 (2026-08-27): an otherwise-excellent answer reported
 *   "Total quantity: 275 units · Pre-tax: ₹14,275 · GST: ₹14,148 · Total: ₹92,748"
 * where pre-tax was really ₹78,600. Three of the four figures were exact, and the
 * Paid/Pending split reconciled to the total precisely — which is exactly what makes
 * this dangerous: one confidently wrong number surrounded by correct ones is the
 * hardest error class for a user to catch, and PRD §2.1 names it as the adoption
 * blocker ("a wrong formula that *looks* right propagates silently into a filing").
 *
 * The error was machine-checkable and unchecked: the answer violated its own
 * arithmetic by a factor of five. This module does not attempt to verify figures
 * against the sheet — it only catches an answer that contradicts ITSELF, which is
 * cheap, deterministic, and has no false-positive risk from unrelated numbers.
 */

/** Tolerance for rounding/display differences (0.5%, plus a small absolute floor). */
const RELATIVE_TOLERANCE = 0.005;
const ABSOLUTE_TOLERANCE = 1;

export type ConsistencyIssue = {
  kind: 'subtotal_mismatch';
  /** Human-readable labels of the operands, as they appeared in the answer. */
  parts: string[];
  expected: number;
  stated: number;
};

/** Parse "₹14,275" / "14,275.50" / "1.4 lakh"-free plain numerics into a number. */
function parseAmount(raw: string): number | null {
  const cleaned = raw.replace(/[^0-9.]/g, '');
  if (!cleaned || cleaned === '.') return null;
  const value = Number(cleaned);
  return Number.isFinite(value) ? value : null;
}

type LabelledAmount = { label: string; value: number };

/**
 * Pull "Label: <amount>" pairs out of an answer. Handles the bullet/middot layout
 * the model actually produces, plus plain "Label: value" lines.
 */
export function extractLabelledAmounts(answer: string): LabelledAmount[] {
  const results: LabelledAmount[] = [];
  // Label may contain spaces, %, and hyphens; amount may carry a currency symbol.
  const re = /([A-Za-z][A-Za-z %()\-]{2,40}?)\s*[:=]\s*([₹$€£]?\s*[\d,]+(?:\.\d+)?)/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(answer)) !== null) {
    const label = match[1].replace(/^[-•·*\s]+/, '').trim();
    const value = parseAmount(match[2]);
    if (value !== null) {
      results.push({ label, value });
    }
  }
  return results;
}

function findByPattern(
  amounts: LabelledAmount[],
  pattern: RegExp,
): LabelledAmount | undefined {
  return amounts.find((a) => pattern.test(a.label));
}

function withinTolerance(expected: number, stated: number): boolean {
  const diff = Math.abs(expected - stated);
  return diff <= Math.max(ABSOLUTE_TOLERANCE, Math.abs(expected) * RELATIVE_TOLERANCE);
}

const PRE_TAX = /\b(pre[\s-]?tax|taxable|net|subtotal|sub[\s-]?total|before tax)\b/i;
const TAX = /\b(gst|tax|vat)\b/i;
const TOTAL = /\b(total|gross|grand total)\b/i;
/**
 * Counts and quantities are not money and must never be treated as the grand total.
 * "Total quantity: 275 units" sits in the same line as the amounts it would corrupt.
 */
const NON_MONETARY = /\b(quantity|qty|units?|records?|rows?|invoices?|count|items?)\b/i;

/**
 * Check that a stated total equals stated pre-tax + tax, when all three appear.
 * Returns null when the answer does not state enough to check.
 */
export function checkAnswerConsistency(answer: string): ConsistencyIssue | null {
  if (!answer || !answer.trim()) return null;
  const amounts = extractLabelledAmounts(answer);
  const monetary = amounts.filter((a) => !NON_MONETARY.test(a.label));
  if (monetary.length < 3) return null;

  const preTax = findByPattern(monetary, PRE_TAX);
  // "Total" must not also match the pre-tax or tax label (e.g. "Total tax").
  const total = monetary.find(
    (a) => TOTAL.test(a.label) && !PRE_TAX.test(a.label) && !TAX.test(a.label),
  );
  const tax = monetary.find((a) => TAX.test(a.label) && !PRE_TAX.test(a.label));

  if (!preTax || !tax || !total) return null;
  if (withinTolerance(preTax.value + tax.value, total.value)) return null;

  return {
    kind: 'subtotal_mismatch',
    parts: [preTax.label, tax.label, total.label],
    expected: preTax.value + tax.value,
    stated: total.value,
  };
}

/**
 * Note appended to an answer that contradicts itself. Deliberately does NOT guess
 * which figure is wrong — only that they cannot all be right — so the user is
 * pointed at the discrepancy rather than handed a second unverified number.
 */
export function describeConsistencyIssue(issue: ConsistencyIssue): string {
  const [preTaxLabel, taxLabel, totalLabel] = issue.parts;
  return (
    `\n\n⚠️ **Please verify these figures.** ${preTaxLabel} + ${taxLabel} comes to ` +
    `${issue.expected.toLocaleString('en-IN')}, but ${totalLabel} is stated as ` +
    `${issue.stated.toLocaleString('en-IN')} — these cannot all be correct, so at ` +
    `least one is wrong. Ask me to recompute any figure you need to rely on.`
  );
}

/** Append a warning to an answer when it contradicts itself; otherwise return as-is. */
export function annotateAnswerConsistency(answer: string): {
  answer: string;
  issue: ConsistencyIssue | null;
} {
  const issue = checkAnswerConsistency(answer);
  if (!issue) return { answer, issue: null };
  return { answer: `${answer}${describeConsistencyIssue(issue)}`, issue };
}
