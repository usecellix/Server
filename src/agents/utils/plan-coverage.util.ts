import { PlanPhase, PlannerOutput, SubTask, WorkbookContext } from '../types/agent.types';

/**
 * Deterministic plan-coverage safety nets (TASKS.md #229, #230).
 *
 * Both close gaps where a prompt rule tells the Planner what to do and nothing
 * checks that it did. Observed live on the 12-month booking-ledger prompt
 * (2026-09-10): the coarse pass emitted a `repeatFor` phase for all 12 months,
 * the expansion pass returned a subtask for January only, and the finished
 * build had two sheets (Main, January) plus 36 `#REF!` cells on Main pointing
 * at February–December. January's dropdowns also pointed at a `Lists` sheet no
 * subtask ever created.
 */

const normalizeSheet = (name: string): string => name.trim().toLowerCase();

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Replace a repeat entry's name with another's, including the space-stripped
 * form used in table names ("Jan 2026" → tblJan2026). A name starting/ending
 * in a digit is guarded so "Unit 1" never rewrites inside "Unit 10".
 */
export function substituteRepeatEntry(text: string, from: string, to: string): string {
  const variants: Array<[string, string]> = [[from, to]];
  const compactFrom = from.replace(/\s+/g, '');
  if (compactFrom !== from) variants.push([compactFrom, to.replace(/\s+/g, '')]);

  let out = text;
  for (const [f, t] of variants) {
    const pattern = new RegExp(
      `${/^\d/.test(f) ? '(?<!\\d)' : ''}${escapeRegExp(f)}${/\d$/.test(f) ? '(?!\\d)' : ''}`,
      'g',
    );
    out = out.replace(pattern, () => t);
  }
  return out;
}

export interface RepeatCoverageResult {
  subtasks: SubTask[];
  /** Entries that had no subtask and were cloned from the template entry. */
  filled: string[];
  /** Template entry the clones were derived from, when any were filled. */
  template?: string;
}

/**
 * A `repeatFor` phase must end up with subtasks for EVERY entry. The expansion
 * prompt asks for one group per entry, but that is probabilistic — when the
 * model returns only the first entry's group, clone that group once per
 * missing entry with the entry name substituted. Cloning (rather than another
 * LLM call per month) is sound precisely because `repeatFor` means "the SAME
 * structure applied once per entry".
 *
 * Entries named inside the template's own descriptions are treated as already
 * covered — a single subtask that explicitly handles "January through
 * December" must not be cloned eleven more times.
 */
export function ensureRepeatForCoverage(phase: PlanPhase, subtasks: SubTask[]): RepeatCoverageResult {
  const entries = phase.repeatFor ?? [];
  if (entries.length < 2 || subtasks.length === 0) return { subtasks, filled: [] };

  const covered = new Set(subtasks.map((s) => normalizeSheet(s.targetSheet)));
  const template =
    entries.find((e) => normalizeSheet(e) === normalizeSheet(phase.targetSheet) && covered.has(normalizeSheet(e))) ??
    entries.find((e) => covered.has(normalizeSheet(e)));
  if (!template) return { subtasks, filled: [] };

  const templateGroup = subtasks.filter((s) => normalizeSheet(s.targetSheet) === normalizeSheet(template));
  const templateText = templateGroup.map((s) => s.description).join('\n');
  const missing = entries.filter(
    (entry) =>
      !covered.has(normalizeSheet(entry)) &&
      !new RegExp(`(^|[^A-Za-z0-9])${escapeRegExp(entry)}([^A-Za-z0-9]|$)`).test(templateText),
  );
  if (missing.length === 0) return { subtasks, filled: [] };

  const usedIds = new Set(subtasks.map((s) => s.id));
  const clones: SubTask[] = [];
  for (const entry of missing) {
    const n = entries.indexOf(entry) + 1;
    const idMap = new Map<string, string>();
    for (const s of templateGroup) {
      let id = `${s.id}_r${n}`;
      while (usedIds.has(id)) id = `${id}x`;
      usedIds.add(id);
      idMap.set(s.id, id);
    }
    for (const s of templateGroup) {
      clones.push({
        ...s,
        id: idMap.get(s.id)!,
        targetSheet: entry,
        description: substituteRepeatEntry(s.description, template, entry),
        dependsOn: s.dependsOn.map((dep) => idMap.get(dep) ?? dep),
      });
    }
  }

  return { subtasks: [...subtasks, ...clones], filled: missing, template };
}

/** `Lists!$B$3`, `'Jan 2026'!A1`, `January!G:G` — sheet name before `!` + a cell/column ref. */
const SHEET_REF_PATTERN =
  /(?:'((?:[^']|'')+)'|([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*))!\$?[A-Za-z]{1,3}\$?\d*/g;

export interface ReferencedSheetResult {
  plan: PlannerOutput;
  /** Sheets that were referenced but never created, now given a create subtask. */
  added: string[];
}

/**
 * Every sheet a subtask references (`Lists!$B$3:$B$20` in a dropdown source,
 * `January!G:G` in a formula) must either exist already or be created by some
 * subtask. When one is neither, add a subtask that creates it and make every
 * referencing subtask depend on it — otherwise the reference lands in the
 * workbook pointing at nothing.
 */
export function ensureReferencedSheetsPlanned(plan: PlannerOutput, context: WorkbookContext): ReferencedSheetResult {
  if (plan.subtasks.length === 0) return { plan, added: [] };

  const known = new Set<string>([
    ...context.sheets.map((s) => normalizeSheet(s.name)),
    ...plan.subtasks.map((s) => normalizeSheet(s.targetSheet)),
  ]);
  // A sheet named in plain prose somewhere ("create sheets January through
  // December") is planned by that subtask even if it is not its targetSheet —
  // only a name that appears SOLELY inside `Sheet!ref` references is orphaned.
  const prose = plan.subtasks.map((s) => s.description.replace(SHEET_REF_PATTERN, ' ')).join('\n');
  const mentionedInProse = (name: string) =>
    new RegExp(`(^|[^A-Za-z0-9_])${escapeRegExp(name)}([^A-Za-z0-9_]|$)`, 'i').test(prose);

  const missing = new Map<
    string,
    { name: string; refs: string[]; referencedBy: Set<string>; backsDropdown: boolean }
  >();
  for (const subtask of plan.subtasks) {
    for (const match of subtask.description.matchAll(SHEET_REF_PATTERN)) {
      const name = (match[1]?.replace(/''/g, "'") ?? match[2] ?? '').trim();
      if (!name || known.has(normalizeSheet(name)) || mentionedInProse(name)) continue;

      const key = normalizeSheet(name);
      const entry = missing.get(key) ?? { name, refs: [], referencedBy: new Set<string>(), backsDropdown: false };
      const before = subtask.description.slice(Math.max(0, (match.index ?? 0) - 60), match.index);
      const hint = before.split(/[,;:\n]/).pop()?.replace(/[(\s]+$/, '').trim();
      const ref = hint ? `${match[0]} (${hint})` : match[0];
      if (!entry.refs.includes(ref)) entry.refs.push(ref);
      if (/dropdown|validation|list/i.test(before)) entry.backsDropdown = true;
      entry.referencedBy.add(subtask.id);
      missing.set(key, entry);
    }
  }

  if (missing.size === 0) return { plan, added: [] };

  const addedSubtasks: SubTask[] = [];
  const dependencyFor = new Map<string, string[]>();
  let n = 0;
  for (const { name, refs, referencedBy, backsDropdown } of missing.values()) {
    n += 1;
    const id = `auto_sheet_${n}`;
    const content = backsDropdown
      ? `Write a short header above each referenced range and fill the range with the values that dropdown ` +
        `should offer; where the real values are unknown, seed a single "Unassigned" placeholder and say so in ` +
        `the summary rather than inventing values. Hide the sheet last — it only backs dropdowns.`
      : `Create the sheet only — do not invent data for it; the referencing formulas read whatever the user enters later.`;
    addedSubtasks.push({
      id,
      targetSheet: name,
      dependsOn: [],
      estimatedActions: Math.min(2 + refs.length, 8),
      description:
        `Create the supporting sheet '${name}' — other steps in this plan reference it but no step creates it. ` +
        `Referenced ranges: ${refs.join('; ')}. ${content}`,
    });
    for (const subtaskId of referencedBy) {
      dependencyFor.set(subtaskId, [...(dependencyFor.get(subtaskId) ?? []), id]);
    }
  }

  const subtasks = plan.subtasks.map((s) =>
    dependencyFor.has(s.id) ? { ...s, dependsOn: [...s.dependsOn, ...dependencyFor.get(s.id)!] } : s,
  );

  return {
    plan: { ...plan, subtasks: [...addedSubtasks, ...subtasks] },
    added: addedSubtasks.map((s) => s.targetSheet),
  };
}
