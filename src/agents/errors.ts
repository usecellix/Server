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
