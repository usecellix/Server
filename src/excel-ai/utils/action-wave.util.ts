import { SheetActionPayload, SheetActionType } from '../types/sheet-actions.types';

/**
 * Sheet-creation actions. When one of these and a write targeting a *different*
 * sheet land in the same batch, the write can only succeed once the create has
 * actually run — a large multi-sheet build (e.g. "make a sheet per month, then
 * fill each in") produces exactly this shape. RichActionEngine already applies
 * actions in array order within one batch, so ordering alone made this work
 * once handleWorksheetAction stopped resolving worksheets it doesn't handle.
 * Splitting into accept waves is therefore not a correctness requirement any
 * more — it is a reviewability one: a 133-action wall of text is hard to
 * review, so users get "review the new sheets" separately from "review what
 * goes in them."
 */
const STRUCTURAL_WAVE_ACTION_TYPES = new Set<SheetActionType>([
  'ADD_SHEET',
  'CREATE_SHEET',
  'COPY_SHEET',
  'RENAME_SHEET',
  'CREATE_TABLE',
  'DEFINE_NAMED_RANGE',
]);

const SHEET_CREATE_TYPES = new Set<SheetActionType>(['ADD_SHEET', 'CREATE_SHEET', 'COPY_SHEET']);

export interface ActionWave {
  actions: SheetActionPayload[];
  /** Accept-card label, e.g. "Create 12 sheets" / "Fill in 121 changes". */
  label: string;
}

/**
 * Split a flattened, already-verified action list into staged accept waves:
 * sheet-creation first, everything else second. Only splits when BOTH
 * categories are non-empty — a pure-write request (the common case) returns
 * a single wave identical to the input, so nothing changes for it.
 */
export function splitIntoActionWaves(actions: SheetActionPayload[]): ActionWave[] {
  const structural = actions.filter((action) => STRUCTURAL_WAVE_ACTION_TYPES.has(action.type));
  const rest = actions.filter((action) => !STRUCTURAL_WAVE_ACTION_TYPES.has(action.type));

  if (structural.length === 0 || rest.length === 0) {
    return [{ actions, label: describeWave(actions) }];
  }

  return [
    { actions: structural, label: describeWave(structural) },
    { actions: rest, label: describeWave(rest) },
  ];
}

function describeWave(actions: SheetActionPayload[]): string {
  const sheetCreates = actions.filter((a) => SHEET_CREATE_TYPES.has(a.type)).length;
  if (sheetCreates > 0 && sheetCreates === actions.length) {
    return `Create ${sheetCreates} sheet${sheetCreates === 1 ? '' : 's'}`;
  }
  return `${actions.length} change${actions.length === 1 ? '' : 's'} ready for review`;
}
