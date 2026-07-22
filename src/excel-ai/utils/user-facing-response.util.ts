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
  /\b(Tier\s*[0-3]|single-action|no verification|Direct Change|Planner|Executor|Verifier|CONDITIONAL_FORMAT|SET_FORMULA|WRITE_TABLE|openai\/)/i;

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
  let text = answer
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

  const bullets =
    actionLines.length >= 2 ? actionLines.map((l) => (l.endsWith('.') ? l.slice(0, -1) : l)) : undefined;

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
