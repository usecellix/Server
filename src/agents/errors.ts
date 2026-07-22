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
