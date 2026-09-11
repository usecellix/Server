import { SheetActionPayload, SheetActionType } from '../types/sheet-actions.types';

/**
 * Split a build into sequential, individually-acceptable STEPS — TASKS.md #160.
 *
 * ## Why this exists, and why it reverses TASKS.md #141
 *
 * #141 collapsed staged accept waves into a single Accept because the two-wave
 * split had a real failure: a user who accepted only "create 13 sheets" and
 * stopped was left with thirteen empty tabs and **nothing telling them the
 * build was half-applied**. That diagnosis was right; the remedy was too blunt.
 * The problem was never staging — it was that staging was *invisible*.
 *
 * Shortcut's transcript shows the opposite trade paying off: six named steps,
 * each applied to the workbook the moment it is accepted, so the sheet visibly
 * becomes real as you go. That is a materially better experience for a long
 * build than one 300-action button, and it is what this restores — with the
 * two things #141's version lacked:
 *
 *   1. every step reports `stepIndex` / `stepTotal`, so "you are not finished"
 *      is stated rather than inferred;
 *   2. an Accept All path, so the staging is never a tax on someone who just
 *      wants the whole thing.
 *
 * ## Why these phases, in this order
 *
 * The order is not cosmetic — it is what makes each step safe to apply on its
 * own, which is the whole precondition for staging:
 *
 *   CREATE  sheets must exist before anything writes into them
 *   CONTENT values before the formulas that reference them
 *   FORMULA cross-sheet references resolve once their targets hold content
 *   FORMAT  number formats and bands apply to written cells
 *   LAYOUT  autofit/freeze measure content, so they follow it
 *   CHART   charts read a source range that must already be populated
 *
 * This is the same order the single-wave build already produced implicitly
 * (structural hoisted to the front, presentation appended at the end); naming
 * the phases just makes it reviewable.
 *
 * Deliberately **structural, never vocabulary-driven** — it keys off action
 * type alone, so it works for any prompt in any domain, not the ledger build it
 * was developed against (the TASKS.md #158 lesson).
 */

/** Ordered phases. Every action type lands in exactly one. */
const PHASES: Array<{ key: string; label: string; types: SheetActionType[] }> = [
  {
    key: 'create',
    label: 'Create sheets',
    types: ['ADD_SHEET', 'CREATE_SHEET', 'COPY_SHEET', 'RENAME_SHEET'],
  },
  {
    key: 'content',
    label: 'Write content',
    types: ['SET_CELL', 'BATCH_SET', 'ADD_ROW', 'INSERT_ROW', 'WRITE_TABLE', 'SET_RANGE_VALUES'],
  },
  {
    key: 'formula',
    label: 'Add formulas',
    types: ['SET_FORMULA', 'FILL_DOWN', 'FILL_RIGHT', 'AGGREGATE_TABLE'],
  },
  {
    key: 'format',
    label: 'Apply formatting',
    types: [
      'FORMAT_RANGE', 'HIGHLIGHT_CELL', 'SET_ROW_HEIGHT', 'SET_COLUMN_WIDTH',
      'MERGE_CELLS', 'UNMERGE_CELLS', 'CONDITIONAL_FORMAT', 'FORMAT_MATCHING_ROWS',
      'SET_SHEET_COLOR',
    ],
  },
  {
    key: 'layout',
    label: 'Finish layout',
    types: [
      'AUTOFIT_COLUMNS', 'FREEZE_PANES', 'UNFREEZE_PANES', 'AUTO_FILTER',
      'CREATE_TABLE', 'DEFINE_NAMED_RANGE', 'HIDE_SHEET', 'SHOW_SHEET',
      'HIDE_ROW', 'HIDE_COLUMN', 'PROTECT_SHEET',
    ],
  },
  {
    key: 'chart',
    label: 'Add charts',
    types: ['CREATE_CHART', 'UPDATE_CHART'],
  },
];

const PHASE_OF = new Map<SheetActionType, number>();
PHASES.forEach((phase, index) => {
  for (const type of phase.types) PHASE_OF.set(type, index);
});

/**
 * Anything unclassified goes with CONTENT — the middle of the order, after
 * sheets exist and before formatting. A new action type therefore degrades to a
 * sensible position rather than being dropped, which is the failure mode the
 * `action-catalog.ts` exhaustiveness work exists to prevent elsewhere.
 */
const DEFAULT_PHASE = 1;

const SHEET_CREATE_TYPES = new Set<SheetActionType>(['ADD_SHEET', 'CREATE_SHEET', 'COPY_SHEET']);

/**
 * Below this many actions, staging is noise: a handful of changes is easier to
 * review as one card than as three. Long builds are what benefit from steps.
 */
const MIN_ACTIONS_TO_STAGE = 25;

/** Never present more steps than a person will read through. */
const MAX_STEPS = 6;

export interface ActionWave {
  actions: SheetActionPayload[];
  /** Accept-card label, e.g. "Create 12 sheets" / "Apply formatting". */
  label: string;
  /**
   * Positions of this wave's actions in the ORIGINAL input array.
   *
   * Carried so a caller can map a wave back to the plan subtasks that produced
   * its actions without re-deriving identity from the rewritten action objects
   * (see `wave-intent.util.ts` for why that matching has to be structural).
   * Kept in step with `actions` through every merge below. TASKS.md #167.
   */
  actionIndexes: number[];
}

function phaseIndexOf(action: SheetActionPayload): number {
  return PHASE_OF.get(action.type) ?? DEFAULT_PHASE;
}

/**
 * Order actions into one wave per non-empty phase.
 *
 * Returns a SINGLE wave when the batch is small, or when everything lands in
 * one phase — so a two-cell edit is never presented as a multi-step ceremony.
 */
export function splitIntoActionWaves(actions: SheetActionPayload[]): ActionWave[] {
  if (actions.length === 0) {
    return [{ actions, label: describeWave(actions), actionIndexes: [] }];
  }

  const buckets = PHASES.map(() => [] as number[]);
  actions.forEach((action, index) => buckets[phaseIndexOf(action)].push(index));

  const nonEmpty = buckets
    .map((indexes, index) => ({ indexes, index }))
    .filter(({ indexes }) => indexes.length > 0);

  const pick = (indexes: number[]) => indexes.map((i) => actions[i]);

  // One phase, or a small batch: a single card is the better review unit.
  if (nonEmpty.length <= 1 || actions.length < MIN_ACTIONS_TO_STAGE) {
    const orderedIndexes = nonEmpty.flatMap(({ indexes }) => indexes);
    const ordered = pick(orderedIndexes);
    return [{ actions: ordered, label: describeWave(ordered), actionIndexes: orderedIndexes }];
  }

  const steps: ActionWave[] = nonEmpty.map(({ indexes, index }) => ({
    actions: pick(indexes),
    label: describeStep(PHASES[index].label, pick(indexes)),
    actionIndexes: indexes,
  }));

  // Merge the smallest neighbouring steps until within budget, rather than
  // truncating — every action must remain in exactly one step.
  while (steps.length > MAX_STEPS) {
    let smallest = 0;
    for (let i = 1; i < steps.length - 1; i += 1) {
      if (steps[i].actions.length < steps[smallest].actions.length) smallest = i;
    }
    const mergeInto = smallest === 0 ? 1 : smallest - 1;
    const [a, b] = mergeInto < smallest ? [mergeInto, smallest] : [smallest, mergeInto];
    const mergedActions = [...steps[a].actions, ...steps[b].actions];
    const mergedIndexes = [...steps[a].actionIndexes, ...steps[b].actionIndexes];
    steps.splice(a, 2, {
      actions: mergedActions,
      label: describeWave(mergedActions),
      actionIndexes: mergedIndexes,
    });
  }

  return steps;
}

function describeStep(phaseLabel: string, actions: SheetActionPayload[]): string {
  const sheetCreates = actions.filter((a) => SHEET_CREATE_TYPES.has(a.type)).length;
  if (sheetCreates > 0 && sheetCreates === actions.length) {
    return `Create ${sheetCreates} sheet${sheetCreates === 1 ? '' : 's'}`;
  }
  const sheets = new Set(
    actions.map((a) => String(a.sheetName ?? '').trim()).filter(Boolean),
  );
  return sheets.size > 1 ? `${phaseLabel} on ${sheets.size} sheets` : phaseLabel;
}

function describeWave(actions: SheetActionPayload[]): string {
  const sheetCreates = actions.filter((a) => SHEET_CREATE_TYPES.has(a.type)).length;
  if (sheetCreates > 0 && sheetCreates === actions.length) {
    return `Create ${sheetCreates} sheet${sheetCreates === 1 ? '' : 's'}`;
  }
  return `${actions.length} change${actions.length === 1 ? '' : 's'} ready for review`;
}
