/** Max chars for userMessage / raw LLM text in the planner log file. */
export const PLANNER_LOG_TEXT_MAX = 50_000;

export interface PlannerLogInputSummary {
  prompt: string;
  routerAssumption?: string;
  /** Full planner user payload (may be truncated). */
  userMessage: string;
  userMessageTruncated?: boolean;
  historyLength: number;
  sheets: string[];
  activeSheet: string;
  hasPromptContext: boolean;
  /** Included only when PLANNER_LOG_FULL_PROMPTS=true. */
  systemPrompt?: string;
}

export interface PlannerLogOutputSummary {
  raw: string;
  rawTruncated?: boolean;
  parsed: unknown;
  fallback: boolean;
  retried: boolean;
}

export interface PlannerLogEntry {
  ts: string;
  correlationId: string;
  model: string;
  durationMs: number;
  success: boolean;
  error?: string;
  input: PlannerLogInputSummary;
  output: PlannerLogOutputSummary;
}

export function truncateForPlannerLog(
  text: string,
  max = PLANNER_LOG_TEXT_MAX,
): { value: string; truncated: boolean } {
  if (text.length <= max) {
    return { value: text, truncated: false };
  }
  return {
    value: `${text.slice(0, max)}\n…[truncated ${text.length - max} chars]`,
    truncated: true,
  };
}

export function formatPlannerLogLine(entry: PlannerLogEntry): string {
  return `${JSON.stringify(entry)}\n`;
}
