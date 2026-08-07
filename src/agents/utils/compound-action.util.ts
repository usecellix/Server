import { Action, ExecutorOutput, SubTask, WorkbookContext } from '../types/agent.types';
import { buildSortFallbackAction } from './sort-action.util';
import {
  detectSheetDataGenerationIntent,
  extractSheetNameFromPrompt,
} from '../../excel-ai/utils/table-request.util';

export function detectCreateNewSheet(text: string): boolean {
  return /\b(create|add)\s+(?:an?\s+)?(?:(?:new|empty|blank)\s+)*sheet/i.test(text);
}

export function detectSortIntent(text: string): boolean {
  return (
    /\bsort(?:\s+the\s+values?\s+of|\s+(?:the\s+)?(?:sheet\s+)?(?:based\s+on|by|on)|\s+based\s+on|\s+by|\s+on|\s+column\b)/i.test(
      text,
    ) || /\bin\s+(?:ascending|descending)\s+order\b/i.test(text)
  );
}

export function extractSortPhrase(prompt: string): string | undefined {
  const match = /\bsort\b.+/i.exec(prompt);
  return match?.[0]?.trim();
}

function nextUniqueSheetName(base: string, context: WorkbookContext): string {
  const existing = new Set(context.sheets.map((sheet) => sheet.name.toLowerCase()));
  if (!existing.has(base.toLowerCase())) return base;
  for (let i = 2; i <= 99; i += 1) {
    const candidate = `${base} ${i}`;
    if (!existing.has(candidate.toLowerCase())) return candidate;
  }
  return `${base} ${Date.now()}`;
}

export function suggestNewSheetName(prompt: string, context: WorkbookContext): string {
  const fromSubtask = extractSheetNameFromSubtaskDescription(prompt, context);
  if (fromSubtask) return fromSubtask;

  if (/\bsorted\b/i.test(prompt)) return nextUniqueSheetName('Sorted', context);
  if (/\bcgst\b/i.test(prompt)) return nextUniqueSheetName('CGST Sorted', context);
  return nextUniqueSheetName('Sheet2', context);
}

export function extractQuotedSheetName(description: string): string | undefined {
  const double = /"([^"]+)"/.exec(description)?.[1]?.trim();
  if (double) return double;
  const single = /'([^']+)'/.exec(description)?.[1]?.trim();
  return single || undefined;
}

/**
 * Pull a sheet name from planner/executor subtask text like:
 *   Create sheet 'paid paid purchases' if it doesn't exist
 *   Create a new sheet called "Paid Purchases"
 * Falls back to extractSheetNameFromPrompt (named/called).
 */
export function extractSheetNameFromSubtaskDescription(
  description: string,
  context: WorkbookContext,
): string | undefined {
  const quoted = extractQuotedSheetName(description);
  if (quoted) return nextUniqueSheetName(quoted, context);

  const sheetClause =
    /\b(?:create|add)\s+(?:an?\s+)?(?:(?:new|empty|blank)\s+)*sheet\s+(?:named\s+|called\s+)?["']?([A-Za-z][A-Za-z0-9 _-]{0,30}?)["']?(?=\s+if\b|\s+when\b|\s+and\b|\s+then\b|\s+with\b|\s+to\b|\s+for\b|\s*,|\s*$)/i.exec(
      description,
    )?.[1]?.trim();
  if (sheetClause && !/^(if|when|with|and|then|for)$/i.test(sheetClause)) {
    return nextUniqueSheetName(sheetClause, context);
  }

  const fromPrompt = extractSheetNameFromPrompt(description);
  return fromPrompt ? nextUniqueSheetName(fromPrompt, context) : undefined;
}

/** Drop generic Sheet2/SheetN creates when the batch already creates the real dest sheet. */
export function pruneSpuriousAddSheetActions(actions: Action[]): Action[] {
  if (actions.length < 2) return actions;

  const namedCreates = new Set(
    actions
      .filter((a) => a.type === 'ADD_SHEET' || a.type === 'CREATE_SHEET')
      .map((a) => String(a.name ?? a.sheetName ?? '').trim().toLowerCase())
      .filter((n) => n && !/^sheet\d*$/i.test(n)),
  );

  if (namedCreates.size === 0) return actions;

  const destSheets = new Set(
    actions
      .filter(
        (a) =>
          a.type === 'COPY_FILTERED_RANGE' ||
          a.type === 'MOVE_RANGE' ||
          a.type === 'AGGREGATE_TABLE',
      )
      .map((a) => String((a as { destSheet?: string }).destSheet ?? '').trim().toLowerCase())
      .filter(Boolean),
  );

  return actions.filter((action) => {
    if (action.type !== 'ADD_SHEET' && action.type !== 'CREATE_SHEET') return true;
    const name = String(action.name ?? action.sheetName ?? '').trim();
    if (!name) return false;
    // Drop Excel-default placeholders when a real named sheet is also being created
    // (or is the COPY/MOVE destination).
    if (/^sheet\d*$/i.test(name)) {
      if (namedCreates.size > 0) return false;
      if (destSheets.size > 0) return false;
    }
    return true;
  });
}

export function detectEmptySheetIntent(text: string): boolean {
  return /\b(empty|blank)\b/i.test(text);
}

export function detectCopySheetIntent(text: string): boolean {
  return /\b(as\s+a\s+copy|copy\s+of|duplicate|clone)\b/i.test(text);
}

/** When copyFrom is omitted, creates a blank sheet via worksheets.add(name). */
export function buildCreateSheetAction(newSheetName: string, copyFrom?: string): Action {
  const action: Action = {
    type: 'ADD_SHEET',
    name: newSheetName,
  };
  if (copyFrom) {
    action.copyFrom = copyFrom;
  }
  return action;
}

function shouldCopyActiveSheetForCreate(desc: string): boolean {
  if (detectEmptySheetIntent(desc)) return false;
  if (detectCopySheetIntent(desc)) return true;
  if (detectSortIntent(desc)) return true;
  return false;
}

function cloneSheetForSort(context: WorkbookContext, newSheetName: string, copyFrom: string): WorkbookContext {
  const source = context.sheets.find((sheet) => sheet.name === copyFrom);
  if (!source || context.sheets.some((sheet) => sheet.name === newSheetName)) {
    return context;
  }

  return {
    ...context,
    sheets: [...context.sheets, { ...source, name: newSheetName }],
    activeSheetName: newSheetName,
  };
}

export function buildCompoundCreateAndSortActions(
  text: string,
  context: WorkbookContext,
  subtask: SubTask,
): ExecutorOutput | null {
  const newSheetName = suggestNewSheetName(text, context);
  const copyFrom = context.activeSheetName;
  const sortContext = cloneSheetForSort(context, newSheetName, copyFrom);
  const sortSubtask: SubTask = {
    ...subtask,
    description: text,
    targetSheet: newSheetName,
  };
  const sortAction = buildSortFallbackAction(sortSubtask, sortContext);
  if (!sortAction) return null;

  sortAction.sheetName = newSheetName;
  return {
    subtaskId: subtask.id,
    actions: [buildCreateSheetAction(newSheetName, copyFrom), sortAction],
    isDone: true,
  };
}

export function buildDeterministicSubtaskActions(
  subtask: SubTask,
  context: WorkbookContext,
): ExecutorOutput | null {
  const desc = subtask.description;

  if (detectCreateNewSheet(desc) && detectSortIntent(desc)) {
    return buildCompoundCreateAndSortActions(desc, context, subtask);
  }

  if (detectCreateNewSheet(desc) && detectSheetDataGenerationIntent(desc)) {
    if (!detectCopySheetIntent(desc) && !detectSortIntent(desc)) {
      return null;
    }
  }

  if (detectCreateNewSheet(desc)) {
    const newName =
      extractSheetNameFromSubtaskDescription(desc, context) ??
      suggestNewSheetName(desc, context);
    // Blank sheet for copy/filter destinations — never clone the active sheet
    // (that fills A1 and trips the overwrite guard on the subsequent COPY).
    const copyFrom = shouldCopyActiveSheetForCreate(desc)
      ? context.activeSheetName
      : undefined;
    return {
      subtaskId: subtask.id,
      actions: [buildCreateSheetAction(newName, copyFrom)],
      isDone: true,
    };
  }

  if (detectSortIntent(desc)) {
    const sortAction = buildSortFallbackAction(subtask, context);
    if (!sortAction) return null;
    return {
      subtaskId: subtask.id,
      actions: [sortAction],
      isDone: true,
    };
  }

  return null;
}

export function maybeMarkSubtaskComplete(
  result: ExecutorOutput,
  subtask: SubTask,
): ExecutorOutput {
  if (result.isDone || result.actions.length === 0) return result;

  const hasSheetCreate = result.actions.some(
    (action) =>
      action.type === 'ADD_SHEET' || action.type === 'CREATE_SHEET' || action.type === 'COPY_SHEET',
  );
  const hasSort = result.actions.some((action) => action.type === 'SORT_RANGE');

  if (detectCreateNewSheet(subtask.description) && detectSortIntent(subtask.description) && hasSheetCreate && hasSort) {
    return { ...result, isDone: true };
  }
  if (detectSortIntent(subtask.description) && hasSort && !detectCreateNewSheet(subtask.description)) {
    return { ...result, isDone: true };
  }
  if (detectCreateNewSheet(subtask.description) && !detectSortIntent(subtask.description) && hasSheetCreate) {
    return { ...result, isDone: true };
  }

  return result;
}

export function buildCompoundFallbackSubtasks(
  prompt: string,
  context: WorkbookContext,
): SubTask[] | null {
  if (!detectCreateNewSheet(prompt) || !detectSortIntent(prompt)) return null;

  const activeSheet = context.activeSheetName || 'Sheet1';
  const newSheetName = suggestNewSheetName(prompt, context);
  const sortPhrase = extractSortPhrase(prompt) ?? 'Sort data';

  return [
    {
      id: 's1',
      description: `Create new sheet "${newSheetName}" as a copy of "${activeSheet}"`,
      targetSheet: activeSheet,
      dependsOn: [],
      estimatedActions: 1,
    },
    {
      id: 's2',
      description: `${sortPhrase} on sheet "${newSheetName}"`,
      targetSheet: newSheetName,
      dependsOn: ['s1'],
      estimatedActions: 1,
    },
  ];
}
