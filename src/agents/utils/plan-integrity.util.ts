import { PlannerOutput, SubTask, WorkbookContext } from '../types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 6 — catch a plan that is already
 * provably wrong BEFORE a single Executor call is spent on it.
 *
 * Every existing coverage net (`ensureTargetSheetsCreated`,
 * `ensureReferencedSheetsPlanned`, `ensureRepeatForCoverage`) reasons from what
 * the PLAN says. None of them reasons from what the USER asked for, and that
 * is precisely the gap TASKS.md #285 fell through: the request named twelve
 * months, the months phase expanded to nothing, and no net noticed because
 * nothing left in the plan referenced the missing sheets either. The user
 * found out from an empty workbook several minutes later.
 *
 * This check is deliberately cheap and deterministic — no model call — so it
 * can run on every plan. Where a gap can be closed by cloning a sibling that
 * IS present, it is closed; where it cannot, it is reported honestly rather
 * than discovered from the finished workbook.
 */

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

/** "all months", "every month", "each month", "monthly", "12 months". */
const ASKS_FOR_ALL_MONTHS =
  /\b(?:all|every|each|twelve|12)\s+(?:the\s+)?months?\b|\bmonthly\b|\bmonths?\s+(?:in|of)\s+(?:a|the|one)\s+year\b/i;

const normalize = (name: string): string => name.trim().toLowerCase();

export type PlanIntegrityKind =
  | 'empty-plan'
  | 'missing-repeated-entity'
  | 'dangling-dependency'
  | 'uncreated-target-sheet';

export interface PlanIntegrityViolation {
  kind: PlanIntegrityKind;
  detail: string;
  /** True when nothing downstream can recover from this. */
  fatal: boolean;
}

export interface PlanIntegrityResult {
  violations: PlanIntegrityViolation[];
  /** The plan, with any deterministically-closable gap already closed. */
  plan: PlannerOutput;
  /** Entities that were missing and have now been synthesized. */
  repaired: string[];
}

/**
 * Which calendar months the request implies but the plan never targets.
 *
 * Only fires when the prompt actually asks for a full year — a request naming
 * one month is not missing the other eleven. Month detection is the concrete,
 * unambiguous case this class of prompt keeps hitting; the same shape extends
 * to other repeated entities when there is evidence they need it, rather than
 * guessed at now.
 */
function missingMonths(prompt: string, plan: PlannerOutput): string[] {
  const namedInPrompt = MONTHS.filter((month) =>
    new RegExp(`\\b${month}\\b`, 'i').test(prompt),
  );
  const wantsWholeYear = ASKS_FOR_ALL_MONTHS.test(prompt) || namedInPrompt.length >= 6;
  if (!wantsWholeYear) return [];

  const planned = new Set(plan.subtasks.map((subtask) => normalize(subtask.targetSheet ?? '')));
  return MONTHS.filter((month) => !planned.has(normalize(month)));
}

/** Clones a month subtask that IS planned, re-targeted at a missing one. */
function cloneMonthSubtask(template: SubTask, month: string, usedIds: Set<string>): SubTask {
  let id = `integrity_${month.toLowerCase()}`;
  while (usedIds.has(id)) id = `${id}x`;
  usedIds.add(id);

  const swap = (text: string): string =>
    text.replace(new RegExp(`\\b${template.targetSheet}\\b`, 'gi'), month);

  return {
    ...template,
    id,
    targetSheet: month,
    description: swap(template.description),
    // The template's own dependencies still apply (Lists, etc.); its
    // month-specific ones do not exist for this clone.
    dependsOn: template.dependsOn.filter((dep) => !dep.includes(template.targetSheet.toLowerCase())),
  };
}

export function checkPlanIntegrity(input: {
  prompt: string;
  plan: PlannerOutput;
  context: WorkbookContext;
}): PlanIntegrityResult {
  const { prompt, context } = input;
  let plan = input.plan;
  const violations: PlanIntegrityViolation[] = [];
  const repaired: string[] = [];

  if (plan.subtasks.length === 0) {
    violations.push({
      kind: 'empty-plan',
      detail: 'The plan contains no steps at all.',
      fatal: true,
    });
    return { violations, plan, repaired };
  }

  // 1. Entities the REQUEST named that the plan never targets — the #285 gap.
  const missing = missingMonths(prompt, plan);
  if (missing.length > 0) {
    const template = plan.subtasks.find((subtask) =>
      MONTHS.some((month) => normalize(month) === normalize(subtask.targetSheet ?? '')),
    );

    if (template) {
      const usedIds = new Set(plan.subtasks.map((subtask) => subtask.id));
      const clones = missing.map((month) => cloneMonthSubtask(template, month, usedIds));
      plan = { ...plan, subtasks: [...plan.subtasks, ...clones] };
      repaired.push(...missing);
      violations.push({
        kind: 'missing-repeated-entity',
        detail:
          `The request asks for every month but the plan covered only ` +
          `${12 - missing.length} of 12 — added ${missing.join(', ')} from the ones it did plan.`,
        fatal: false,
      });
    } else {
      violations.push({
        kind: 'missing-repeated-entity',
        detail:
          `The request asks for every month but the plan targets none of them ` +
          `(${missing.length} missing), and there is no month step to model them on.`,
        fatal: true,
      });
    }
  }

  // 2. Structural checks. Non-fatal: later nets already repair these, and this
  //    is here so a violation is on the record rather than silently patched.
  const ids = new Set(plan.subtasks.map((subtask) => subtask.id));
  const dangling = plan.subtasks.flatMap((subtask) =>
    subtask.dependsOn.filter((dep) => !ids.has(dep)).map((dep) => `${subtask.id} -> ${dep}`),
  );
  if (dangling.length > 0) {
    violations.push({
      kind: 'dangling-dependency',
      detail: `Dependencies that name no step in this plan: ${dangling.join(', ')}.`,
      fatal: false,
    });
  }

  const existing = new Set(context.sheets.map((sheet) => normalize(sheet.name)));
  const creates = new Set(
    plan.subtasks
      .filter((subtask) => /\bcreate\b|\badd\b|\bnew sheet\b/i.test(subtask.description))
      .map((subtask) => normalize(subtask.targetSheet ?? '')),
  );
  const uncreated = [
    ...new Set(
      plan.subtasks
        .map((subtask) => subtask.targetSheet?.trim())
        .filter((sheet): sheet is string => Boolean(sheet))
        .filter((sheet) => !existing.has(normalize(sheet)) && !creates.has(normalize(sheet))),
    ),
  ];
  if (uncreated.length > 0) {
    violations.push({
      kind: 'uncreated-target-sheet',
      detail: `Steps target sheets nothing creates: ${uncreated.join(', ')}.`,
      fatal: false,
    });
  }

  return { violations, plan, repaired };
}
