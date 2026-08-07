import { SheetAction } from '../types/sheet-actions.types';
import { resolveActionWriteTarget } from './overwrite-confirmation.util';

/**
 * Spec 18 / Spec 13 chart-identity extension + Spec 21 overwrite-refinement:
 * structured record of successful writes / charts so follow-ups can resolve
 * "the current" ranges and recognize refinements of this conversation's own edits.
 */
export interface TurnActionRecord {
  actionType: string;
  sheetName: string;
  /** Spec 21 — range this turn wrote into (sheet-qualified when known). */
  affectedRange?: string;
  /** Spec 21 — named column for SET_MATCHING_ROWS-style writes. */
  targetColumn?: string;
  turnIndex?: number;
  sourceRange?: string;
  sourceSheetName?: string;
  destStartCell?: string;
  destSheet?: string;
  chartId?: string;
  chartType?: string;
  groupByColumn?: string;
}

const CHART_TABLE_TYPES = new Set(['CREATE_CHART', 'UPDATE_CHART', 'AGGREGATE_TABLE']);

const WRITE_TRACKED_TYPES = new Set([
  'SET_CELL',
  'SET_FORMULA',
  'BATCH_SET',
  'WRITE_TABLE',
  'FILL_DOWN',
  'FILL_RIGHT',
  'AUTO_FILL',
  'SET_MATCHING_ROWS',
  'COPY_FILTERED_RANGE',
  'MOVE_RANGE',
  'AGGREGATE_TABLE',
]);

const TRACKED_TYPES = new Set([...CHART_TABLE_TYPES, ...WRITE_TRACKED_TYPES]);

export function extractTurnActionRecords(actions: unknown[] | undefined | null): TurnActionRecord[] {
  if (!Array.isArray(actions)) return [];
  const records: TurnActionRecord[] = [];

  for (const raw of actions) {
    if (!raw || typeof raw !== 'object') continue;
    const action = raw as SheetAction;
    if (!TRACKED_TYPES.has(action.type)) continue;

    if (action.type === 'CREATE_CHART' || action.type === 'UPDATE_CHART') {
      const sheetName = String(
        action.sheetName ?? action.sourceSheetName ?? action.sourceSheet ?? '',
      ).trim();
      if (!sheetName && !action.chartId) continue;
      const sourceRange = optionalString(action.sourceRange ?? action.range);
      records.push({
        actionType: action.type,
        sheetName: sheetName || 'Sheet1',
        affectedRange: sourceRange
          ? qualify(sheetName || 'Sheet1', sourceRange)
          : undefined,
        sourceRange,
        sourceSheetName: optionalString(action.sourceSheetName ?? action.sourceSheet),
        destStartCell: optionalString(action.destCell ?? action.destStartCell),
        chartId: optionalString(action.chartId),
        chartType: optionalString(action.chartType),
      });
      continue;
    }

    if (action.type === 'AGGREGATE_TABLE') {
      const writeTarget = resolveActionWriteTarget(action as never);
      const sheetName = String(action.destSheet ?? action.sheetName ?? '').trim();
      records.push({
        actionType: 'AGGREGATE_TABLE',
        sheetName: sheetName || String(action.sourceSheet ?? 'Sheet1'),
        affectedRange: writeTarget?.affectedRange,
        sourceRange: optionalString(action.sourceRange ?? action.range),
        sourceSheetName: optionalString(action.sourceSheet),
        destStartCell: optionalString(action.destStartCell),
        destSheet: optionalString(action.destSheet),
        groupByColumn: optionalString(action.groupByColumn),
      });
      continue;
    }

    // Spec 21 write tracking
    const writeTarget = resolveActionWriteTarget(action as never);
    if (!writeTarget) continue;
    records.push({
      actionType: action.type,
      sheetName: writeTarget.sheetName,
      affectedRange: writeTarget.affectedRange,
      targetColumn: writeTarget.targetColumn,
      sourceRange: optionalString(action.sourceRange ?? action.range),
      destSheet: optionalString(action.destSheet),
      destStartCell: optionalString(action.destStartCell),
    });
  }

  return records;
}

type StoredTurnActionRecord = Omit<TurnActionRecord, 'actionType'> & {
  actionType: string;
};

function coerceTurnActionRecord(raw: StoredTurnActionRecord): TurnActionRecord | null {
  if (!TRACKED_TYPES.has(raw.actionType) || typeof raw.sheetName !== 'string') return null;
  return {
    ...raw,
    actionType: raw.actionType,
  };
}

export function collectRecentTurnActionRecords(
  messages: Array<{
    role?: string;
    metadata?: { actions?: unknown[]; turnActionRecords?: StoredTurnActionRecord[] };
  }>,
  limit = 8,
): TurnActionRecord[] {
  const collected: TurnActionRecord[] = [];

  for (let i = messages.length - 1; i >= 0 && collected.length < limit; i -= 1) {
    const msg = messages[i];
    if (msg.role !== 'assistant') continue;

    const fromMeta = Array.isArray(msg.metadata?.turnActionRecords)
      ? msg.metadata!.turnActionRecords!
          .map(coerceTurnActionRecord)
          .filter((r): r is TurnActionRecord => r != null)
      : [];
    const fromActions = extractTurnActionRecords(msg.metadata?.actions);
    const batch = fromMeta.length > 0 ? fromMeta : fromActions;

    for (const record of batch) {
      collected.push({ ...record, turnIndex: i });
      if (collected.length >= limit) break;
    }
  }

  return collected.reverse();
}

/** Follow-ups that should prefer structured prior chart/table ranges. */
export function referencesPriorChartOrTable(message: string): boolean {
  return (
    /\b(the\s+current|that\s+chart|this\s+chart|the\s+same\s+data|along\s+with\s+the\s+current|same\s+(?:range|table|data|chart)|existing\s+chart)\b/i.test(
      message,
    ) || /\b(too|also)\b/i.test(message)
  );
}

export function formatTurnActionRecordsForExecutor(records: TurnActionRecord[]): string {
  if (records.length === 0) return '';

  const lines = records.map((r, index) => {
    const bits = [`#${index + 1} ${r.actionType}`];
    if (r.chartId) bits.push(`chartId=${r.chartId}`);
    if (r.chartType) bits.push(`chartType=${r.chartType}`);
    if (r.sourceSheetName) bits.push(`sourceSheet=${r.sourceSheetName}`);
    if (r.sourceRange) bits.push(`sourceRange=${r.sourceRange}`);
    if (r.affectedRange) bits.push(`affectedRange=${r.affectedRange}`);
    if (r.targetColumn) bits.push(`targetColumn=${r.targetColumn}`);
    if (r.destSheet) bits.push(`destSheet=${r.destSheet}`);
    if (r.destStartCell) bits.push(`destStartCell=${r.destStartCell}`);
    if (r.groupByColumn) bits.push(`groupBy=${r.groupByColumn}`);
    bits.push(`sheet=${r.sheetName}`);
    return `- ${bits.join(', ')}`;
  });

  return [
    'Prior turn actions (authoritative — reuse ranges for "the current" / "same data"; refinements of affectedRange may overwrite with confirmation):',
    ...lines,
  ].join('\n');
}

/**
 * Prefer the most recent CREATE_CHART / AGGREGATE_TABLE source range when the
 * user references "the current" data.
 */
export function resolvePriorSourceRange(records: TurnActionRecord[]): {
  sourceRange?: string;
  sourceSheetName?: string;
  chartId?: string;
} | null {
  for (let i = records.length - 1; i >= 0; i -= 1) {
    const r = records[i];
    if (r.actionType === 'CREATE_CHART' && r.sourceRange) {
      return {
        sourceRange: r.sourceRange,
        sourceSheetName: r.sourceSheetName ?? r.sheetName,
        chartId: r.chartId,
      };
    }
    if (r.actionType === 'AGGREGATE_TABLE' && (r.destStartCell || r.sourceRange)) {
      // Follow-up charts usually plot the aggregate output, not the raw source.
      if (r.destSheet && r.destStartCell) {
        return {
          sourceRange: undefined,
          sourceSheetName: r.destSheet,
          chartId: undefined,
        };
      }
      if (r.sourceRange) {
        return {
          sourceRange: r.sourceRange,
          sourceSheetName: r.sourceSheetName ?? r.sheetName,
        };
      }
    }
  }
  return null;
}

function optionalString(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  return trimmed || undefined;
}

function qualify(sheetName: string, local: string): string {
  const sheet = sheetName.trim();
  const range = local.trim();
  if (!range) return '';
  return sheet ? `${sheet}!${range}` : range;
}
