export class StepRetryExhaustedError extends Error {
  public readonly step: unknown;
  public readonly attempts: number;

  constructor(message: string, context: { step: unknown; attempts: number }) {
    super(message);
    this.name = 'StepRetryExhaustedError';
    this.step = context.step;
    this.attempts = context.attempts;
  }
}

/** Planner LLM returned empty/unparseable output after all retries — never emit a stub plan. */
export class PlannerExhaustedError extends Error {
  public readonly originalMessage: string;

  constructor(
    message: string,
    context: { originalMessage: string },
  ) {
    super(message);
    this.name = 'PlannerExhaustedError';
    this.originalMessage = context.originalMessage;
  }
}

export const PLANNER_EXHAUSTED_USER_MESSAGE =
  "I had trouble planning this request — it may be too complex for one step. Try breaking it into smaller requests (e.g. first 'create a Dashboard sheet', then 'add a chart of...').";

/**
 * Planner declined to make the LLM call at all because its estimated cost
 * exceeds the per-call cost cap — a distinct failure mode from
 * `PlannerExhaustedError` (the model tried and produced nothing usable).
 * Kept separate so cost refusals don't blur into `ALERT
 * reasoning_token_exhaustion`-style exhaustion telemetry, and so the
 * user-facing message can point at the actual cause (request/context too
 * large) instead of "too complex for one step."
 */
export class PlannerCostCapExceededError extends Error {
  public readonly originalMessage: string;
  public readonly estimatedCostUsd: number;
  public readonly costCapUsd: number;
  public readonly promptTokens: number;
  public readonly maxTokens: number;

  constructor(
    message: string,
    context: {
      originalMessage: string;
      estimatedCostUsd: number;
      costCapUsd: number;
      promptTokens: number;
      maxTokens: number;
    },
  ) {
    super(message);
    this.name = 'PlannerCostCapExceededError';
    this.originalMessage = context.originalMessage;
    this.estimatedCostUsd = context.estimatedCostUsd;
    this.costCapUsd = context.costCapUsd;
    this.promptTokens = context.promptTokens;
    this.maxTokens = context.maxTokens;
  }
}

export const PLANNER_COST_CAP_USER_MESSAGE =
  "This request's context is too large to plan safely in one call. Try focusing on fewer sheets, trimming the conversation history, or splitting the request into smaller pieces.";
