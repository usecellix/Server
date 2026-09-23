/**
 * Is this string something Excel would actually accept as a sheet name?
 *
 * TASKS.md #290 — the sheet-reference regexes in `plan-coverage.util.ts` and
 * `reconcile.util.ts` match a quoted name before a `!`. A formula that builds
 * its reference dynamically defeats that completely:
 *
 *   INDIRECT("'"&TEXT(DATE(2026,ROW(INDIRECT("1:12")),1),"mmmm")&"'!A:A")
 *
 * The text between the two apostrophes is a formula fragment, not a name, and
 * a live run took it at face value — the coverage net created a real sheet
 * called `"&TEXT(DATE(2026,ROW(INDIRECT("1:12")),1),"mmmm")&"` and put it in
 * the workbook.
 *
 * The rules below are Excel's own documented ones, not heuristics, which is
 * what makes this safe to apply: legitimate names people really use ("P&L",
 * "Q1 (Draft)", "Jan 2026") all pass, while the fragment above fails twice
 * over — it is 48 characters and contains a colon.
 */

/** Excel forbids these outright in a sheet name. */
const ILLEGAL_SHEET_CHARS = /[\\/?*[\]:]/;

/** Excel's own limit. */
const MAX_SHEET_NAME_LENGTH = 31;

export function isPlausibleSheetName(candidate: string): boolean {
  const name = candidate.trim();
  if (!name) return false;
  if (name.length > MAX_SHEET_NAME_LENGTH) return false;
  if (ILLEGAL_SHEET_CHARS.test(name)) return false;
  // Excel also rejects a name that starts or ends with an apostrophe.
  if (name.startsWith("'") || name.endsWith("'")) return false;
  return true;
}
