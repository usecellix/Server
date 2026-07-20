import { AssistantMode, SheetActionPayload, SheetActionType } from '../types/sheet-actions.types';

/**
 * Control-only action types that do not mutate workbook data. Everything else
 * is treated as a write and is stripped in read-only modes (ask / plan).
 */
const NON_MUTATING_ACTION_TYPES = new Set<SheetActionType>(['CLARIFY', 'CHECKPOINT']);

export function isWriteAction(action: SheetActionPayload): boolean {
  return !NON_MUTATING_ACTION_TYPES.has(action.type);
}

export function modeIsReadOnly(mode: AssistantMode | 'act' | undefined): boolean {
  const normalized = normalizeAssistantMode(mode);
  return normalized === 'ask' || normalized === 'plan';
}

/** Normalize API aliases (`act`) and default omitted mode to `action`. */
export function normalizeAssistantMode(mode?: string): AssistantMode {
  if (mode === 'ask' || mode === 'plan') {
    return mode;
  }
  return 'action';
}

export interface StripResult {
  actions: SheetActionPayload[];
  removedCount: number;
}

/**
 * Defense-in-depth: in ask/plan modes, remove any write actions the LLM may
 * have produced so that no silent edits can ever reach the add-in.
 */
export function stripWriteActions(actions: SheetActionPayload[]): StripResult {
  const kept = actions.filter((action) => !isWriteAction(action));
  return { actions: kept, removedCount: actions.length - kept.length };
}
