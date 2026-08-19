/**
 * Never invent display formats (especially dates). Only apply a numberFormat when:
 * - the user named that format string, or
 * - it already appears on the target column in workbook context (re-apply existing).
 */
import type { Action, PlannerOutput, WorkbookContext } from '../types/agent.types';

const EXPLICIT_FORMAT_HINT_RE =
  /\b(dd[-\/]?mm[-\/]?yyyy|mm[-\/]?dd[-\/]?yyyy|yyyy[-\/]?mm[-\/]?dd|d\/m\/yyyy|m\/d\/yyyy|dd\.mm\.yyyy)\b/i;

/** Well-known product defaults — never treat as "original" unless the sheet already uses them. */
const PRODUCT_DEFAULT_DATE_FORMATS = new Set(
  ['dd-mm-yyyy', 'dd/mm/yyyy', 'dd.mm.yyyy'].map((f) => f.toLowerCase()),
);

export function isNumberFormatCode(text: string): boolean {
  const t = text.trim().toLowerCase();
  if (!t || t === 'general') return false;
  // Excel format patterns always include y/m/d/# or currency/symbol mixes
  return /[ymd#0₹$%]/.test(t) || EXPLICIT_FORMAT_HINT_RE.test(t);
}

export function isDateNumberFormat(format: string): boolean {
  const lower = format.trim().toLowerCase();
  if (!lower || lower === 'general') return false;
  return (
    (lower.includes('y') || lower.includes('d')) &&
    lower.includes('m') &&
    !lower.includes('#,') // not Indian number alone
  );
}

export function extractUserNamedNumberFormats(message: string): string[] {
  const found: string[] = [];
  const quoted = message.matchAll(/["'`]([^"'`]{2,40})["'`]/g);
  for (const m of quoted) {
    const v = m[1]!.trim();
    if (isNumberFormatCode(v)) found.push(v);
  }
  // Unquoted common tokens
  const tokenMatch = message.match(EXPLICIT_FORMAT_HINT_RE);
  if (tokenMatch) found.push(tokenMatch[0]);
  return found;
}

export function userNamedFormat(message: string, format: string): boolean {
  const named = extractUserNamedNumberFormats(message);
  const target = format.trim().toLowerCase();
  return named.some((n) => n.trim().toLowerCase() === target);
}

/** Vague restore / reformat without a concrete format code. */
export function isVagueDateFormatRequest(message: string): boolean {
  const lower = message.toLowerCase();
  if (!/\bdate\b|\bdates\b|\bformat\b/.test(lower)) return false;
  if (extractUserNamedNumberFormats(message).length > 0) return false;
  return (
    /\b(original|back|restore|revert|reset|previous|before)\b/.test(lower) ||
    /\b(change|fix|apply|set|fix)\b.+\b(date|dates).+\bformat\b/.test(lower) ||
    /\bdate\b.+\bformat\b/.test(lower)
  );
}

function getSheet(context: WorkbookContext, sheetName?: string) {
  const name = sheetName ?? context.activeSheetName;
  return (
    context.sheets.find((s) => s.name === name) ??
    context.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase())
  );
}

/** Majority non-General format for a column from loaded numberFormats rows. */
export function dominantColumnNumberFormat(
  context: WorkbookContext,
  sheetName: string | undefined,
  col: number,
): string | null {
  const sheet = getSheet(context, sheetName);
  if (!sheet?.numberFormats?.length) return null;

  const counts = new Map<string, number>();
  for (let r = 0; r < sheet.numberFormats.length; r += 1) {
    // Prefer data rows for "existing" format; still count header if it's a real format
    const raw = sheet.numberFormats[r]?.[col];
    if (raw == null) continue;
    const fmt = String(raw).trim();
    if (!fmt || /^general$/i.test(fmt)) continue;
    counts.set(fmt, (counts.get(fmt) ?? 0) + 1);
  }
  if (counts.size === 0) return null;
  let best: string | null = null;
  let bestN = 0;
  for (const [fmt, n] of counts) {
    if (n > bestN) {
      best = fmt;
      bestN = n;
    }
  }
  return best;
}

/**
 * Rewrite or drop FORMAT_RANGE numberFormat when the model invents a code
 * the user did not name and the column does not already use.
 */
export function rebindFormatRangeNumberFormats(
  actions: Action[],
  opts: {
    userPrompt: string;
    subtaskDescription?: string;
    context: WorkbookContext;
  },
): { actions: Action[]; droppedInvented: boolean } {
  const prompt = `${opts.userPrompt}\n${opts.subtaskDescription ?? ''}`;
  let droppedInvented = false;

  const next = actions
    .map((action) => {
      if (action.type !== 'FORMAT_RANGE') return action;
      const numberFormat = action.format?.numberFormat;
      if (typeof numberFormat !== 'string' || !numberFormat.trim()) return action;

      if (userNamedFormat(prompt, numberFormat)) {
        return action;
      }

      const col = typeof action.col === 'number' ? action.col : 0;
      const colCount = typeof action.colCount === 'number' && action.colCount > 0 ? action.colCount : 1;
      // Single-col and multi-col: only allow when every targeted col already uses this format
      // or the invented format is replaced by dominant for the first col when rebinding a date col.
      let allowed = true;
      for (let c = col; c < col + colCount; c += 1) {
        const dominant = dominantColumnNumberFormat(opts.context, action.sheetName, c);
        if (dominant && dominant.trim().toLowerCase() === numberFormat.trim().toLowerCase()) {
          continue;
        }
        allowed = false;
        break;
      }

      if (allowed) return action;

      // Rebind date-like invents to the dominant existing format when present
      if (isDateNumberFormat(numberFormat) || PRODUCT_DEFAULT_DATE_FORMATS.has(numberFormat.toLowerCase())) {
        const dominant = dominantColumnNumberFormat(opts.context, action.sheetName, col);
        if (dominant && isDateNumberFormat(dominant)) {
          return {
            ...action,
            format: { ...action.format, numberFormat: dominant },
          };
        }
        // No existing format to re-apply — drop the numberFormat field
        const { numberFormat: _drop, ...restFormat } = action.format ?? {};
        droppedInvented = true;
        if (Object.keys(restFormat).length === 0) {
          return null; // drop pure invent-only format action
        }
        return { ...action, format: restFormat };
      }

      // Non-date invent (e.g. random currency) without user naming: drop numberFormat only
      const { numberFormat: _drop, ...restFormat } = action.format ?? {};
      droppedInvented = true;
      if (Object.keys(restFormat).length === 0) {
        return null;
      }
      return { ...action, format: restFormat };
    })
    .filter((a): a is Action => a !== null);

  return { actions: next, droppedInvented };
}

/**
 * Planner guard: refuse plans that assume product-default date formats without user naming them.
 */
export function ensureNumberFormatPlanSafety(prompt: string, plan: PlannerOutput): PlannerOutput {
  if (plan.clarificationsNeeded.length > 0) return plan;

  const assumesInvented =
    plan.subtasks.some((s) =>
      /\b(assumed|assume)\b.*\b(dd-mm-yyyy|dd\/mm\/yyyy|date format)\b|\bdd-mm-yyyy\b|\bdd\/mm\/yyyy\b/i.test(
        s.description,
      ),
    ) && extractUserNamedNumberFormats(prompt).length === 0;

  const vagueRestore =
    isVagueDateFormatRequest(prompt) &&
    plan.subtasks.some(
      (s) =>
        /\b(FORMAT_RANGE|date format|numberFormat|format the Date)\b/i.test(s.description) &&
        !/\b(existing|current|already on|re-?apply|majority|sample|from (the )?column)\b/i.test(
          s.description,
        ),
    ) &&
    extractUserNamedNumberFormats(prompt).length === 0;

  if (!assumesInvented && !vagueRestore) return plan;

  return {
    ...plan,
    subtasks: [],
    confidence: 'low',
    clarificationsNeeded: [
      ...plan.clarificationsNeeded,
      'Which date (or number) format should I apply? I will not invent one. Prefer naming it (e.g. m/d/yyyy, dd-mm-yyyy, yyyy-mm-dd) or say "use the format already on these cells".',
    ],
    reasoning: `${plan.reasoning} [preserve formats: refuse invented numberFormat]`.trim(),
  };
}
