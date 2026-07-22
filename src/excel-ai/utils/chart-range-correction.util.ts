import { Action, VerifierOutput } from '../../agents/types/agent.types';
import { normalizeChartColorScheme } from '../../agents/utils/chart-color-scheme.util';
import {
  referencesPriorChartOrTable,
  resolvePriorSourceRange,
  TurnActionRecord,
} from '../utils/turn-action-history.util';

const A1_RANGE_RE = /\b([A-Za-z]{1,3}\d{1,7}:[A-Za-z]{1,3}\d{1,7})\b/;
const QUOTED_A1_RANGE_RE = /['"`]([A-Za-z]{1,3}\d{1,7}:[A-Za-z]{1,3}\d{1,7})['"`]/;

/** Follow-ups like "create a bar graph too" also imply prior chart/table context. */
export function referencesPriorChartOrTableOrAlso(message: string): boolean {
  return (
    referencesPriorChartOrTable(message) ||
    /\b(too|also)\b/i.test(message)
  );
}

/**
 * Prefer structured prior CREATE_CHART sourceRange when the user is clearly
 * continuing from a prior chart/table turn.
 */
export function applyPriorSourceRangeToChartActions(
  actions: Action[],
  priorTurnActions: TurnActionRecord[] | WorkbookPriorActions | undefined,
  message: string,
): { actions: Action[]; applied: boolean; sourceRange?: string } {
  if (!actions.length || !priorTurnActions?.length) {
    return { actions, applied: false };
  }
  if (!referencesPriorChartOrTableOrAlso(message)) {
    return { actions, applied: false };
  }

  const prior = resolvePriorSourceRange(priorTurnActions as TurnActionRecord[]);
  if (!prior?.sourceRange) {
    return { actions, applied: false };
  }

  let applied = false;
  const next = actions.map((action) => {
    if (action.type !== 'CREATE_CHART' && action.type !== 'UPDATE_CHART') {
      return action;
    }
    if (action.sourceRange === prior.sourceRange) {
      return action;
    }
    applied = true;
    return {
      ...action,
      sourceRange: prior.sourceRange,
      ...(prior.sourceSheetName
        ? { sourceSheetName: prior.sourceSheetName, sheetName: action.sheetName ?? prior.sourceSheetName }
        : {}),
    };
  });

  return { actions: next, applied, sourceRange: prior.sourceRange };
}

type WorkbookPriorActions = Array<{
  actionType: string;
  sheetName: string;
  sourceRange?: string;
  sourceSheetName?: string;
  destStartCell?: string;
  destSheet?: string;
  chartId?: string;
  chartType?: string;
  groupByColumn?: string;
}>;

/** Pull a concrete A1 suggestion out of verifier suggestion fields (not descriptions). */
export function extractSuggestedSourceRange(verifierResult: VerifierOutput): string | undefined {
  const suggestionTexts: string[] = [];
  for (const issue of verifierResult.issues ?? []) {
    if (issue.suggestion?.trim()) suggestionTexts.push(issue.suggestion);
  }
  for (const sub of verifierResult.subtaskResults ?? []) {
    for (const issue of sub.issues ?? []) {
      if (issue.suggestion?.trim()) suggestionTexts.push(issue.suggestion);
    }
  }

  for (const text of suggestionTexts) {
    const quoted = QUOTED_A1_RANGE_RE.exec(text);
    if (quoted?.[1]) return quoted[1].toUpperCase();
  }
  for (const text of suggestionTexts) {
    const bare = A1_RANGE_RE.exec(text);
    if (bare?.[1]) return bare[1].toUpperCase();
  }

  // Fallback: "Try: A1:B10" style phrases in top-level feedback only
  const feedback = verifierResult.feedback ?? '';
  const tryMatch =
    /(?:try|e\.g\.|for example|use)\s*[:\s]*['"`]?([A-Za-z]{1,3}\d{1,7}:[A-Za-z]{1,3}\d{1,7})/i.exec(
      feedback,
    );
  if (tryMatch?.[1]) return tryMatch[1].toUpperCase();

  return undefined;
}

export function patchChartSourceRanges(
  actions: Action[],
  sourceRange: string,
): { actions: Action[]; patched: boolean } {
  let patched = false;
  const next = actions.map((action) => {
    if (action.type !== 'CREATE_CHART' && action.type !== 'UPDATE_CHART') {
      return action;
    }
    if (action.sourceRange === sourceRange) {
      return action;
    }
    // UPDATE_CHART often has no sourceRange — only patch when present or CREATE_CHART
    if (action.type === 'UPDATE_CHART' && !action.sourceRange) {
      return action;
    }
    patched = true;
    return { ...action, sourceRange };
  });
  return { actions: next, patched };
}

/**
 * If the user asked for a named color and chart actions lack colorScheme,
 * stamp it on CREATE_CHART so the sole LLM retry is not wasted on color alone.
 */
export function ensureRequestedChartColorScheme(
  message: string,
  actions: Action[],
): { actions: Action[]; applied: boolean } {
  const match = /\b(green|red|blue|orange|purple|yellow|grey|gray|blueGrey|bluegrey)\b/i.exec(
    message,
  );
  if (!match?.[1]) return { actions, applied: false };

  const scheme = normalizeChartColorScheme(match[1]);
  if (!scheme) return { actions, applied: false };

  const alreadyColored = actions.some(
    (a) =>
      (a.type === 'CREATE_CHART' || a.type === 'UPDATE_CHART') &&
      a.colorScheme === scheme,
  );
  if (alreadyColored) return { actions, applied: false };

  let applied = false;
  const next = actions.map((action) => {
    if (action.type !== 'CREATE_CHART') return action;
    if (action.colorScheme) return action;
    applied = true;
    return { ...action, colorScheme: scheme };
  });

  if (applied) return { actions: next, applied };

  // No CREATE_CHART to stamp — append a minimal UPDATE_CHART if we have a chartId
  const create = actions.find((a) => a.type === 'CREATE_CHART');
  const chartId = create && 'chartId' in create ? create.chartId : undefined;
  const sheetName = create && 'sheetName' in create ? create.sheetName : undefined;
  if (!chartId || !sheetName) return { actions, applied: false };

  return {
    actions: [
      ...actions,
      {
        type: 'UPDATE_CHART' as const,
        sheetName: String(sheetName),
        chartId: String(chartId),
        colorScheme: scheme,
      },
    ],
    applied: true,
  };
}
