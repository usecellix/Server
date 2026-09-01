import { CellChange } from '../../audit/types/change-set.types';
import { SheetAction } from '../types/sheet-actions.types';

export interface UserFacingSummary {
  contextLine?: string;
  headline: string;
  bullets?: string[];
  supportingDetail?: string;
}

export interface ResponseInternalDetails {
  tier?: number;
  model?: string;
  processingLabel?: string;
  reasoning?: string;
  assumption?: string;
  rawActionSummary?: string;
  legacyExplanation?: string;
}

export interface BuildUserFacingSummaryInput {
  answer?: string;
  actions: SheetAction[];
  changes?: CellChange[];
  assumption?: string;
  activeSheetName?: string;
  /**
   * The plan's own subtasks, when this response came from Tier 3. When present
   * these become the card's bullets, because they describe INTENT ("Create
   * sheet and set A1:J1 headers - 12 sheets") rather than mechanics ("Freeze
   * Panes on June"). Absent for Tier 0-2, which fall back to the action
   * rollup. TASKS.md #149.
   */
  planSubtasks?: PlanIntent[];
}

export interface BuildInternalDetailsInput {
  tier?: number;
  model?: string;
  processingLabel?: string;
  reasoning?: string;
  assumption?: string;
  actions: SheetAction[];
  legacyExplanation?: string;
}

/** Strings that must never appear in the default user-facing headline. */
export const INTERNAL_COPY_MARKERS =
  /\b(Tier\s*[0-3]|single-action|no verification|Direct Change|Planner|Executor|Verifier|CONDITIONAL_FORMAT|FORMAT_MATCHING_ROWS|findMatchingRowOffsets|hasHeaders\s*:|SET_FORMULA|WRITE_TABLE|openai\/)/i;

/** Spec 24: full answers must not leak action-type / validation stack fragments. */
export function sanitizeAnswerForUser(answer: string): string {
  if (!answer) return answer;
  let text = answer;
  if (
    /findMatchingRowOffsets|FORMAT_MATCHING_ROWS\s*:|hasHeaders\s*:\s*true|Spreadsheet update failed/i.test(
      text,
    )
  ) {
    return "I couldn't apply that formatting. Please try again or describe the range differently.";
  }
  // Drop lines that look like internal ActionType: message dumps
  text = text
    .split('\n')
    .filter((line) => !/^[A-Z][A-Z0-9_]+\s*:\s*.+/.test(line.trim()))
    .join('\n')
    .trim();
  if (INTERNAL_COPY_MARKERS.test(text) && !text.includes(' ')) {
    return "I couldn't apply that formatting. Please try again or describe the range differently.";
  }
  return text || answer;
}

function colToLetter(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s || 'A';
}

function parseA1(address: string): { col: number; row: number } | null {
  const m = /^([A-Za-z]+)(\d+)$/.exec(address.trim());
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return { col: col - 1, row: parseInt(m[2], 10) };
}

/** Compact range like "A53:C55" or "Purchase Register!A53:C55". */
export function describeRangeCompactly(changes: CellChange[]): string | undefined {
  if (!changes.length) return undefined;

  const bySheet = new Map<string, CellChange[]>();
  for (const c of changes) {
    const sheet = c.sheet || '';
    const list = bySheet.get(sheet) ?? [];
    list.push(c);
    bySheet.set(sheet, list);
  }

  const parts: string[] = [];
  for (const [sheet, sheetChanges] of bySheet) {
    let minCol = Infinity;
    let maxCol = -Infinity;
    let minRow = Infinity;
    let maxRow = -Infinity;
    let parsedAny = false;

    for (const c of sheetChanges) {
      const parsed = parseA1(c.cell);
      if (!parsed) continue;
      parsedAny = true;
      minCol = Math.min(minCol, parsed.col);
      maxCol = Math.max(maxCol, parsed.col);
      minRow = Math.min(minRow, parsed.row);
      maxRow = Math.max(maxRow, parsed.row);
    }

    if (!parsedAny) {
      parts.push(sheet ? `${sheet} (${sheetChanges.length} cells)` : `${sheetChanges.length} cells`);
      continue;
    }

    const start = `${colToLetter(minCol)}${minRow}`;
    const end = `${colToLetter(maxCol)}${maxRow}`;
    const range = start === end ? start : `${start}:${end}`;
    parts.push(sheet ? `${sheet}!${range}` : range);
  }

  return parts.join(', ');
}

export function sanitizeAnswerForHeadline(answer: string): string {
  let text = sanitizeAnswerForUser(answer)
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\*([^*]+)\*/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/^#+\s*/gm, '')
    .trim();

  // Prefer first paragraph / sentence cluster for the lead line.
  const firstPara = text.split(/\n\n+/)[0]?.trim() ?? text;
  text = firstPara.split(/\n/)[0]?.trim() ?? firstPara;

  if (INTERNAL_COPY_MARKERS.test(text)) {
    return '';
  }

  if (text && !/[.!?]$/.test(text)) {
    text = `${text}.`;
  }
  return text;
}

function describeOneAction(action: SheetAction): string {
  const sheet = action.sheetName ? ` on ${action.sheetName}` : '';
  switch (action.type) {
    case 'FORMAT_RANGE':
    case 'FORMAT_MATCHING_ROWS':
      if (action.format?.bold) return `Apply bold formatting${sheet}`;
      if (action.format?.fillColor) return `Highlight matching cells${sheet}`;
      return `Apply formatting${sheet}`;
    case 'SORT_RANGE':
      return `Sort the sheet${sheet}`;
    case 'SET_CELL':
      return `Update cell values${sheet}`;
    case 'SET_FORMULA':
      return `Add formulas${sheet}`;
    case 'WRITE_TABLE':
    case 'BATCH_SET':
      return `Write table data${sheet}`;
    case 'ADD_ROW':
    case 'INSERT_ROW':
      return `Add row(s)${sheet}`;
    case 'DELETE_ROW':
      return `Delete row(s)${sheet}`;
    case 'CREATE_CHART':
      return `Create a chart${sheet}`;
    case 'AGGREGATE_TABLE':
      return `Add a summary below the table${sheet}`;
    default:
      return action.type
        .toLowerCase()
        .split('_')
        .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
        .join(' ')
        .replace(/\bId\b/g, '')
        .trim() + sheet;
  }
}

/**
 * A plan subtask, reduced to what the Accept card needs.
 */
export interface PlanIntent {
  id: string;
  description: string;
  targetSheet: string;
}

/** Longest a single intent bullet may run before it is trimmed. */
const MAX_INTENT_BULLET_CHARS = 150;

function trimIntent(text: string): string {
  const clean = text.replace(/\s+/g, ' ').trim();
  if (clean.length <= MAX_INTENT_BULLET_CHARS) return clean;
  // Cut on a word boundary so a truncated formula list does not end mid-token.
  const cut = clean.slice(0, MAX_INTENT_BULLET_CHARS);
  const lastSpace = cut.lastIndexOf(' ');
  return `${(lastSpace > 60 ? cut.slice(0, lastSpace) : cut).trimEnd()}...`;
}

const REGEXP_SPECIALS = new Set([
  '.', '*', '+', '?', '^', '$', '{', '}', '(', ')', '|', '[', ']', '/',
  String.fromCharCode(92),
]);

/**
 * Escape a sheet name for literal use inside a RegExp. Written as an explicit
 * character walk rather than a character-class replace so the escaping is
 * readable and cannot itself be mis-escaped — sheet names are arbitrary user
 * strings and routinely contain `(`, `)`, `.` and `-`.
 */
function escapeForRegExp(value: string): string {
  let out = '';
  for (const ch of value) {
    out += REGEXP_SPECIALS.has(ch) ? String.fromCharCode(92) + ch : ch;
  }
  return out;
}

/** Strip every mention of `sheet` (quoted or bare) from `text`. */
function stripSheetName(text: string, sheet: string): string {
  if (!sheet) return text.replace(/\s+/g, ' ').trim();
  const quote = String.fromCharCode(96);
  const pattern = '[' + String.fromCharCode(39) + '"' + quote + ']?' + escapeForRegExp(sheet) + '[' + String.fromCharCode(39) + '"' + quote + ']?';
  return text
    .replace(new RegExp(pattern, 'gi'), ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Group key for a subtask: its description with its own target sheet removed,
 * so twelve structurally identical month subtasks collapse to one entry.
 *
 * Deliberately generic - it keys off `targetSheet`, never a hardcoded month
 * list, so it works equally for "a sheet per region", "per client", "per
 * department". Repeated shape is what long prompts almost always produce, and
 * it is exactly what makes their Accept cards unreadable today.
 */
function intentGroupKey(subtask: PlanIntent): string {
  return stripSheetName(subtask.description, subtask.targetSheet?.trim() ?? '').toLowerCase();
}

/** Human list: "January, February, March +9 more". */
function describeSheetGroup(sheets: string[]): string {
  const shown = sheets.slice(0, 3).join(', ');
  const rest = sheets.length - 3;
  return rest > 0 ? `${shown} +${rest} more` : shown;
}

/**
 * Turn a plan into a short list of intent statements - what the build DOES.
 *
 * This is the semantic half of the Accept card. The Planner already writes the
 * right sentences; they were streamed as transient `status` events and then
 * thrown away, leaving the card to enumerate ~40 mechanical actions instead.
 * Repeated-shape subtasks are grouped and counted rather than listed, which is
 * what takes a 19-subtask monthly-ledger plan down to a handful of readable
 * lines. The full action list still lives behind "Show details". TASKS.md #149.
 */
export function summarizePlanIntent(subtasks: PlanIntent[]): string[] {
  if (!subtasks?.length) return [];

  const groups = new Map<string, { first: PlanIntent; sheets: string[] }>();
  for (const subtask of subtasks) {
    if (!subtask?.description?.trim()) continue;
    const key = intentGroupKey(subtask);
    const sheet = subtask.targetSheet?.trim();
    const existing = groups.get(key);
    if (existing) {
      if (sheet && !existing.sheets.includes(sheet)) existing.sheets.push(sheet);
    } else {
      groups.set(key, { first: subtask, sheets: sheet ? [sheet] : [] });
    }
  }

  return [...groups.values()].map(({ first, sheets }) => {
    if (sheets.length <= 1) return trimIntent(first.description);
    // Strip the one sheet the description happens to name, then say how many
    // sheets this actually covers - otherwise the line reads as if it applied
    // to January alone.
    // Trim the description BEFORE appending the count, never after: the
    // "- 12 sheets (January, February, March +9 more)" suffix is the single
    // most informative part of the line, and trimming the assembled string
    // silently ate it for any description already near the cap.
    const generic = trimIntent(stripSheetName(first.description, sheets[0]));
    return `${generic} - ${sheets.length} sheets (${describeSheetGroup(sheets)})`;
  });
}

/**
 * Intent bullets describe distinct build steps, so a long build legitimately
 * needs more of them than the mechanical action rollup does. Overflow past this
 * collapses into a single "+N more steps" line instead of discarding the list.
 */
const MAX_INTENT_BULLETS = 8;

function capIntentLines(lines: string[]): string[] {
  if (lines.length <= MAX_INTENT_BULLETS) return lines;
  const shown = lines.slice(0, MAX_INTENT_BULLETS - 1);
  return [...shown, `+${lines.length - shown.length} more steps`];
}

/**
 * Most bullets a card body may carry before the list stops being a summary and
 * starts being the wall of text the details disclosure exists for. A 13-sheet
 * build produced ~40 distinct lines ("Freeze Panes on January", "Freeze Panes
 * on February", …) and rendered every one of them in the card. TASKS.md #140.
 */
const MAX_SUMMARY_BULLETS = 6;

/**
 * Collapse per-sheet repetition into one line per kind of change.
 *
 * `describeOneAction` appends " on <sheet>", so the same operation repeated
 * across twelve month sheets reads as twelve distinct lines. Group by the
 * sheet-less description and report the sheet *count* instead — "Freeze panes
 * on 13 sheets" says what forty lines said, in one.
 */
export function rollUpActionsForUser(actions: SheetAction[]): string[] {
  const byVerb = new Map<string, Set<string>>();

  for (const action of actions) {
    const verb = describeOneAction({ ...action, sheetName: undefined });
    const sheets = byVerb.get(verb);
    const sheetName = action.sheetName?.trim();
    if (sheets) {
      if (sheetName) sheets.add(sheetName);
    } else {
      byVerb.set(verb, new Set(sheetName ? [sheetName] : []));
    }
  }

  return [...byVerb.entries()].map(([verb, sheets]) => {
    if (sheets.size === 0) return verb;
    if (sheets.size === 1) return `${verb} on ${[...sheets][0]}`;
    return `${verb} on ${sheets.size} sheets`;
  });
}

/** Distinct plain-English lines for actions (used for bullets / fallback headline). */
export function describeActionsForUser(actions: SheetAction[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const action of actions) {
    const line = describeOneAction(action);
    if (seen.has(line)) continue;
    seen.add(line);
    out.push(line);
  }
  return out;
}

function rawActionLine(action: SheetAction): string {
  const bits: string[] = [action.type];
  if (action.sheetName) bits.push(action.sheetName);
  if (action.range) bits.push(action.range);
  else if (action.address) bits.push(action.address);
  else if (typeof action.row === 'number' && typeof action.col === 'number') {
    bits.push(`${colToLetter(action.col)}${action.row + 1}`);
  }
  return bits.join(' ');
}

export function buildUserFacingSummary(input: BuildUserFacingSummaryInput): UserFacingSummary {
  const { answer, actions, changes = [], assumption, activeSheetName } = input;

  const actionLines = describeActionsForUser(actions);
  let headline =
    (answer ? sanitizeAnswerForHeadline(answer) : '') ||
    (actionLines.length === 1
      ? `${actionLines[0]}.`.replace(/\.\.$/, '.')
      : actionLines.length > 1
        ? `I'll make ${actionLines.length} changes to your sheet.`
        : 'Ready to apply changes.');

  if (assumption?.trim()) {
    const a = assumption.trim();
    const alreadyCovered =
      headline.toLowerCase().includes(a.toLowerCase().slice(0, Math.min(40, a.length))) ||
      a.toLowerCase().includes(headline.toLowerCase().slice(0, Math.min(40, headline.length)));
    if (!alreadyCovered) {
      const assumptionSentence = /[.!?]$/.test(a) ? a : `${a}.`;
      headline = `${assumptionSentence} ${headline}`;
    }
  }

  // Final guard — never ship internal tokens in the default headline.
  if (INTERNAL_COPY_MARKERS.test(headline)) {
    headline =
      actionLines.length > 0
        ? `${actionLines[0]}.`.replace(/\.\.$/, '.')
        : 'Ready to apply changes.';
  }

  // The card body summarizes; the details disclosure enumerates.
  //
  // Preference order, most meaningful first:
  //   1. Plan intent  - what the build DOES (Tier 3 only; TASKS.md #149)
  //   2. Rolled-up actions - one line per kind of change (TASKS.md #140)
  //   3. Nothing - headline + meta only, when even the rollup is too long
  const intentLines = summarizePlanIntent(input.planSubtasks ?? []);
  const rolledLines =
    actionLines.length > MAX_SUMMARY_BULLETS ? rollUpActionsForUser(actions) : actionLines;
  // Intent lines get a larger budget than the action rollup, and overflow into
  // a "+N more steps" tail rather than being dropped wholesale. A long build
  // legitimately HAS eight or nine distinct steps; showing them is the point,
  // and an all-or-nothing cap silently produced a card with no bullets at all
  // for exactly the prompts that need them most.
  const candidate =
    intentLines.length > 0 ? capIntentLines(intentLines) : rolledLines;
  // Two or more, always: a lone bullet only restates the headline, which is the
  // pre-existing invariant the 9-cell/1-action regression test pins.
  const withinBudget =
    intentLines.length > 0 ? candidate.length <= MAX_INTENT_BULLETS : candidate.length <= MAX_SUMMARY_BULLETS;
  const bullets =
    candidate.length >= 2 && withinBudget
      ? candidate.map((l) => (l.endsWith('.') ? l.slice(0, -1) : l))
      : undefined;

  let supportingDetail: string | undefined;
  if (changes.length > 0) {
    const range = describeRangeCompactly(changes);
    const cellLabel = `${changes.length} cell${changes.length === 1 ? '' : 's'}`;
    supportingDetail = range ? `${cellLabel}, ${range}` : cellLabel;
  }

  const contextLine = activeSheetName?.trim()
    ? `Working with: ${activeSheetName.trim()}`
    : undefined;

  return {
    contextLine,
    headline,
    bullets,
    supportingDetail,
  };
}

export function buildInternalDetails(input: BuildInternalDetailsInput): ResponseInternalDetails {
  const rawActionSummary =
    input.actions.length > 0
      ? input.actions.map(rawActionLine).join('; ')
      : undefined;

  return {
    tier: input.tier,
    model: input.model,
    processingLabel: input.processingLabel,
    reasoning: input.reasoning,
    assumption: input.assumption,
    rawActionSummary,
    legacyExplanation: input.legacyExplanation,
  };
}

export function tierProcessingLabel(tier: 0 | 1 | 2 | 3, actionHint?: string): string {
  switch (tier) {
    case 0:
      return 'Tier 0 direct resolution — no LLM calls.';
    case 1:
      return `Tier 1 single-action (${actionHint ?? 'ACTION'}) — one LLM call, no verification.`;
    case 2:
      return 'Tier 2 generate-verify: executed and verified (no Planner).';
    case 3:
      return 'Multi-agent pipeline: planned, executed, and verified.';
    default:
      return `Tier ${tier}`;
  }
}
