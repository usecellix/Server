export class LlmRequestError extends Error {
  constructor(
    public readonly status: number,
    message?: string,
  ) {
    super(message ?? `LLM request failed (${status})`);
    this.name = 'LlmRequestError';
  }

  get isRecoverable(): boolean {
    return [402, 429, 500, 502, 503, 504].includes(this.status);
  }
}

/** @deprecated use LlmRequestError */
export const OpenAiRequestError = LlmRequestError;
