import { SubTask } from '../types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 5 — one line telling the user what a
 * big build is about to do, before it starts doing it.
 *
 * Deliberately NOT a question. Every clarifying question this session added to
 * a long build made things worse, because the thing that actually goes wrong
 * is scale, not ambiguity — and a question mid-flow costs a whole extra
 * round-trip through a path that then has to rediscover the original request.
 * This just says what is coming and roughly how long, so a five-minute build
 * reads as working rather than hung.
 */

/** Below this a build is short enough that announcing it is just noise. */
export const PLAN_SUMMARY_MIN_SUBTASKS = 6;

/**
 * Rough wall-clock per subtask that actually costs an LLM call. Measured from
 * this session's own runs (a month "rest" step lands around 20-40s); it is a
 * ballpark shown as "~N min", never a promise.
 */
const SECONDS_PER_LLM_SUBTASK = 30;

export interface PlanSummary {
  sheetCount: number;
  subtaskCount: number;
  /** Subtasks that cost an LLM call — the deterministic ones are ~free. */
  llmSubtaskCount: number;
  estimatedMinutes: number;
  text: string;
}

export function buildPlanSummary(subtasks: SubTask[]): PlanSummary | null {
  if (subtasks.length < PLAN_SUMMARY_MIN_SUBTASKS) return null;

  const sheets = new Set(
    subtasks.map((subtask) => subtask.targetSheet?.trim()).filter((name): name is string => Boolean(name)),
  );
  const llmSubtaskCount = subtasks.filter((subtask) => !subtask.isDeterministicHeaderStep).length;
  const estimatedMinutes = Math.max(
    1,
    Math.round((llmSubtaskCount * SECONDS_PER_LLM_SUBTASK) / 60),
  );

  const sheetWord = sheets.size === 1 ? 'sheet' : 'sheets';
  const text =
    `Building ${sheets.size} ${sheetWord} in ${subtasks.length} steps — roughly ` +
    `${estimatedMinutes} minute${estimatedMinutes === 1 ? '' : 's'}. ` +
    `You can accept each step as it lands, or stop at any point.`;

  return {
    sheetCount: sheets.size,
    subtaskCount: subtasks.length,
    llmSubtaskCount,
    estimatedMinutes,
    text,
  };
}
