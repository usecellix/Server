import { PlanPhase, PlannerOutput, SubTask, WorkbookContext } from '../types/agent.types';
import { isPlausibleSheetName } from './sheet-name.util';

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

/**
 * Accepts `unknown` on purpose — TASKS.md #263. Both callers read names out of
 * LLM-derived plans and wire-built workbook context, and a single entry whose
 * `name` was not a string threw `TypeError: name.trim is not a function` from
 * inside `ensureReferencedSheetsPlanned`, failing the ENTIRE request after the
 * Planner had already run (~90s and a paid call, thrown away). A malformed
 * sheet entry should cost that entry, not the build.
 */
const normalizeSheet = (name: unknown): string =>
  typeof name === 'string' ? name.trim().toLowerCase() : '';

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
      // TASKS.md #290 — a dynamically-built reference (INDIRECT with a
      // computed name) puts a formula fragment where a sheet name should be.
      // Creating a sheet from it is worse than ignoring it: a live run put a
      // sheet literally named `"&TEXT(DATE(...),"mmmm")&"` in the workbook.
      if (!name || !isPlausibleSheetName(name)) continue;
      if (known.has(normalizeSheet(name)) || mentionedInProse(name)) continue;

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

/**
 * Does this description say it CREATES `sheetName`, as opposed to merely
 * writing to it? "Create sheet 'January' (position after Main)" creates
 * January and only mentions Main; "On Main, write title in A1" creates
 * nothing. Both orderings count — "create sheet 'X'" and "create the X sheet".
 *
 * The gap between the sheet keyword and the name is deliberately tight (12
 * chars) so a create for one sheet cannot be read as a create for another
 * sheet named later in the same sentence.
 */
export function describesSheetCreation(description: string, sheetName: string): boolean {
  const raw = typeof sheetName === 'string' ? sheetName.trim() : '';
  const name = escapeRegExp(raw);
  if (!name) return false;

  const quote = `['"‘’“”]?`;
  const verb = String.raw`\b(?:creat\w*|add|adds|adding|insert\w*|make|makes|making|build\w*|set\s+up)\b`;
  const kind = String.raw`\b(?:sheet|worksheet|tab)\b`;

  // "Create sheet 'January'" / "Create the January tab" — the sheet keyword is
  // present, so the name may sit a little away from the verb.
  const withKeyword =
    `${verb}[^.;\\n]{0,40}?(?:${kind}[^.;\\n]{0,12}?${quote}${name}${quote}` +
    `|${quote}${name}${quote}[^.;\\n]{0,12}?${kind})`;

  // "Create January" / "Create the Main" — no sheet keyword at all, which the
  // Planner does use. Kept TIGHT (verb, optional determiner, then the name) so
  // "Create sheet 'January' (position after Main)" is not read as creating
  // Main: only the sheet the verb actually governs counts.
  const bareName =
    `${verb}\\s+(?:(?:a|an|the|new|empty)\\s+)*${quote}${name}${quote}(?![A-Za-z0-9])`;

  return new RegExp(`(?:${withKeyword})|(?:${bareName})`, 'i').test(description);
}

export interface TargetSheetCreationResult {
  plan: PlannerOutput;
  /** Target sheets nothing created, now given a create subtask. */
  added: string[];
}

/**
 * A sheet that subtasks TARGET must also be created by one of them — TASKS.md
 * #262.
 *
 * `ensureReferencedSheetsPlanned` treats every `targetSheet` as already
 * planned, on the assumption that a subtask owning a sheet creates it. Live on
 * the 12-month ledger prompt that assumption broke: the month subtasks said
 * "Create sheet 'January'…" but all five Main subtasks said "On Main, write
 * …". Main was a `targetSheet`, so it counted as known, so no create was ever
 * planned — and every write to it failed at Accept with "The requested
 * resource doesn't exist", which is what made Accept look like it did nothing.
 *
 * Runs after `ensureReferencedSheetsPlanned` so the `auto_sheet_*` creates it
 * adds are seen here as real creates (their descriptions say so) and are not
 * duplicated. A redundant create would be harmless anyway — the client's
 * `handleAddSheet` reuses an existing sheet rather than making "Main 2" — but
 * not adding one when it is needed loses the whole build.
 */
export function ensureTargetSheetsCreated(
  plan: PlannerOutput,
  context: WorkbookContext,
): TargetSheetCreationResult {
  if (plan.subtasks.length === 0) return { plan, added: [] };

  const existing = new Set(context.sheets.map((s) => normalizeSheet(s.name)));

  const targets: string[] = [];
  for (const subtask of plan.subtasks) {
    const name = typeof subtask.targetSheet === 'string' ? subtask.targetSheet.trim() : '';
    if (!name) continue;
    if (!targets.some((t) => normalizeSheet(t) === normalizeSheet(name))) targets.push(name);
  }

  const addedSubtasks: SubTask[] = [];
  const dependencyFor = new Map<string, string>();
  let n = 0;

  for (const name of targets) {
    if (existing.has(normalizeSheet(name))) continue;
    if (plan.subtasks.some((s) => describesSheetCreation(s.description, name))) continue;

    n += 1;
    const id = `auto_create_${n}`;
    addedSubtasks.push({
      id,
      targetSheet: name,
      dependsOn: [],
      estimatedActions: 1,
      description:
        `Create the sheet '${name}' — later steps write to it but no step creates it. ` +
        `Create it empty and do not invent content; the steps depending on this one fill it in.`,
    });
    for (const subtask of plan.subtasks) {
      if (normalizeSheet(subtask.targetSheet ?? '') === normalizeSheet(name)) {
        dependencyFor.set(subtask.id, id);
      }
    }
  }

  if (addedSubtasks.length === 0) return { plan, added: [] };

  const subtasks = plan.subtasks.map((s) =>
    dependencyFor.has(s.id) ? { ...s, dependsOn: [...s.dependsOn, dependencyFor.get(s.id)!] } : s,
  );

  return {
    plan: { ...plan, subtasks: [...addedSubtasks, ...subtasks] },
    added: addedSubtasks.map((s) => s.targetSheet),
  };
}

/**
 * A last-resort plan for a phase whose expansion produced NOTHING — TASKS.md #285.
 *
 * `expandPhase` already retries a truncated/unparseable response once, then
 * gives up and returns an empty plan. A live run showed what that costs: the
 * coarse pass correctly identified three phases (Lists, the 12 month sheets,
 * Main), phase p2's expansion came back empty, and the build shipped as
 * "Step 1 of 3 … ✓ Applied" with all twelve month sheets missing — the
 * dashboard on Main left summing sheets that were never created.
 * `ensureRepeatForCoverage` could not help: it clones a SIBLING subtask, and
 * here there was no sibling to clone.
 *
 * The coarse phase itself still carries everything needed to state the work:
 * its `kind` (what to build) and, for a repeated structure, every `repeatFor`
 * entry. That is enough for one honest subtask per entry — deliberately plain,
 * since this is a recovery path, and the columns/table each sheet actually
 * gets come from the build spec (Phase 1/1.5), not from this text.
 */
export function synthesizeSubtasksForEmptyPhase(phase: PlanPhase): SubTask[] {
  const entries = (phase.repeatFor?.length ? phase.repeatFor : [phase.targetSheet])
    .map((entry) => entry?.trim())
    .filter((entry): entry is string => Boolean(entry));

  return entries.map((entry, index) => ({
    id: `s${index + 1}`,
    targetSheet: entry,
    dependsOn: [],
    estimatedActions: 12,
    description:
      `Create sheet '${entry}' and build it as described: ${phase.kind}` +
      (entries.length > 1 ? ` (this is the '${entry}' one of ${entries.length}).` : '.'),
  }));
}
