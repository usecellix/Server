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
  const fromPrompt = extractSheetNameFromPrompt(prompt);
  if (fromPrompt) return nextUniqueSheetName(fromPrompt, context);

  if (/\bsorted\b/i.test(prompt)) return nextUniqueSheetName('Sorted', context);
  if (/\bcgst\b/i.test(prompt)) return nextUniqueSheetName('CGST Sorted', context);
  return nextUniqueSheetName('Sheet2', context);
}

export function extractQuotedSheetName(description: string): string | undefined {
  const match = /"([^"]+)"/.exec(description);
  return match?.[1]?.trim();
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
    const newName = extractQuotedSheetName(desc) ?? suggestNewSheetName(desc, context);
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
