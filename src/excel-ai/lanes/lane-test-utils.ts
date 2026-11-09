import { LaneContext, LaneResult } from './lane.types';
import { BatchSetOperation } from '../types/sheet-actions.types';

/** Excel serial for a calendar date (what Office.js returns for a date cell). */
export const serial = (y: number, m: number, d: number): number => Math.round(Date.UTC(y, m - 1, d) / 86_400_000) + 25_569;

export function makeCtx(values: unknown[][], over: Partial<LaneContext> = {}, others: Record<string, unknown[][]> = {}): LaneContext {
  return {
    message: '',
    writeAllowed: true,
    sheetName: 'S',
    values,
    headerRowIndex: 0,
    columnMeta: [],
    sheetNames: ['S', ...Object.keys(others)],
    hiddenSheets: [],
    getSheetValues: async (name) => (name === (over.sheetName ?? 'S') ? values : others[name] ?? null),
    now: new Date(Date.UTC(2026, 10, 7)),
    ...over,
  };
}

/** address → value|formula for the first BATCH_SET in a result. */
export function cellsOf(result: LaneResult | null, index = 0): Record<string, unknown> {
  if (!result || result.kind !== 'actions') return {};
  const batch = result.actions.filter((a) => a.type === 'BATCH_SET')[index] as { operations: BatchSetOperation[] } | undefined;
  return Object.fromEntries((batch?.operations ?? []).map((o) => [o.address, o.formula ?? o.value]));
}
