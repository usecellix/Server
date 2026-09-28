import { Action, SubTask } from '../types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 2 — N structurally identical sibling
 * subtasks (12 month sheets) are built by ONE Executor run plus N-1
 * deterministic clones of its accepted actions.
 *
 * A live 12-month build fired 12 simultaneous LLM generations that were all
 * supposed to produce the same structure; a different random subset failed on
 * every run. The plan itself already proves they are the same work — the
 * descriptions differ only by the sheet name — so generating each one
 * separately buys nothing but load and drift.
 */

/** Below this, cloning saves too little to be worth the risk of a wrong merge. */
export const MIN_CLONE_GROUP_SIZE = 3;

export interface CloneGroup {
  template: SubTask;
  clones: SubTask[];
}

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * The name plus, for plain alphabetic names, its 3-letter form — planners
 * abbreviate the table name ('tblJan' for January) and the abbreviation is
 * as much a part of "the sheet's name" as the name itself.
 */
function nameForms(sheetName: string): string[] {
  const name = sheetName.trim();
  const forms = [name];
  if (/^[A-Za-z]{4,}$/.test(name)) forms.push(name.slice(0, 3));
  return forms;
}

/**
 * Replaces each form of `from` with the matching form of `to`. Full name first,
 * so "January" is consumed before its abbreviation "Jan" is looked for. The
 * `(?![a-z])` guard keeps "Mar" from matching inside "Market" while still
 * matching inside "tblMar".
 */
function renameSheetTokens(text: string, from: string, to: string): string {
  const fromForms = nameForms(from);
  const toForms = nameForms(to);
  let out = text;
  fromForms.forEach((form, index) => {
    const replacement = toForms[index] ?? toForms[0];
    if (!form) return;
    out = out.replace(new RegExp(`${escapeRegExp(form)}(?![a-z])`, 'g'), () => replacement);
  });
  return out;
}

/** Same placeholder substitution `normalizedSignature` applies to a description. */
function normalizeSheetTokens(text: string, sheetName: string): string {
  return nameForms(sheetName).reduce(
    (out, form) => out.replace(new RegExp(`${escapeRegExp(form)}(?![a-z])`, 'g'), '\u0001SHEET\u0001'),
    text,
  );
}

function normalizedSignature(subtask: SubTask): string {
  // Placeholder is deliberately not a real word so it cannot collide with content.
  // One placeholder for both forms: "May" has no separate abbreviation, so its
  // 'tblMay' must normalize the same way as 'tblJan' does for January.
  const description = normalizeSheetTokens(subtask.description, subtask.targetSheet);
  // A split subtask (header-table-split.util.ts) names its own helper id from
  // the sheet ("hdr_January") — without normalizing dependsOn the SAME way,
  // every split "rest" subtask would carry a DIFFERENT-looking dependency
  // across months and silently opt back out of cloning entirely.
  const dependsOn = [...subtask.dependsOn]
    .map((id) => normalizeSheetTokens(id, subtask.targetSheet))
    .sort();
  return JSON.stringify([
    description,
    dependsOn,
    subtask.estimatedActions,
    subtask.suggestedActionType ?? null,
    subtask.expectedHeaders ?? null,
  ]);
}

/**
 * Groups subtasks of one wave that are the same work on different sheets.
 * Anything that differs beyond the sheet name (a per-month rule, a different
 * dependency) is left out and simply runs normally.
 */
export function findCloneGroups(wave: SubTask[]): CloneGroup[] {
  const bySignature = new Map<string, SubTask[]>();
  for (const subtask of wave) {
    // Two subtasks on the SAME sheet are sequential steps, never siblings.
    if (!subtask.targetSheet.trim()) continue;
    const key = normalizedSignature(subtask);
    bySignature.set(key, [...(bySignature.get(key) ?? []), subtask]);
  }

  const groups: CloneGroup[] = [];
  for (const members of bySignature.values()) {
    const sheets = new Set(members.map((member) => member.targetSheet.trim().toLowerCase()));
    if (members.length < MIN_CLONE_GROUP_SIZE || sheets.size !== members.length) continue;
    groups.push({ template: members[0], clones: members.slice(1) });
  }
  return groups;
}

function renameDeep(value: unknown, from: string, to: string): unknown {
  if (typeof value === 'string') return renameSheetTokens(value, from, to);
  if (Array.isArray(value)) return value.map((entry) => renameDeep(entry, from, to));
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>).map(([key, entry]) => [
        key,
        renameDeep(entry, from, to),
      ]),
    );
  }
  return value;
}

/**
 * Re-targets a template's accepted actions at another sheet: `sheetName`,
 * sheet-qualified formula references, and table names all follow the rename.
 * References to OTHER sheets (Lists!$B$3) are untouched — they name neither
 * form of the template's sheet.
 */
export function cloneActionsForSheet(
  actions: Action[],
  fromSheet: string,
  toSheet: string,
): Action[] {
  return actions.map((action) => renameDeep(action, fromSheet, toSheet) as Action);
}
