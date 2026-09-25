import { AsyncLocalStorage } from 'node:async_hooks';

/**
 * Per-request attribution for LLM calls. Set once at the HTTP boundary and read
 * by OpenRouterService at the moment a call hits the network, so the ~17 call
 * sites across agents/services never have to thread a prompt id through.
 *
 * Mutable on purpose: some fields are only known partway through a request
 * (the conversationId is minted by the service; a /continue only learns its
 * prompt id after loading the run; route/tier are decided by the router).
 */
export interface LlmUsageContext {
  /** Groups every call made on behalf of one user prompt, across /continue waves. */
  promptId: string;
  userId?: string;
  conversationId?: string;
  route?: string;
  tier?: number;
  /** Last error surfaced to the user in this request, thrown or sent as an SSE `error`. */
  error?: string;
}

const storage = new AsyncLocalStorage<LlmUsageContext>();

export function runWithLlmUsageContext<T>(context: LlmUsageContext, fn: () => T): T {
  return storage.run(context, fn);
}

export function currentLlmUsageContext(): LlmUsageContext | undefined {
  return storage.getStore();
}

export function updateLlmUsageContext(patch: Partial<LlmUsageContext>): void {
  const store = storage.getStore();
  if (!store) return;
  for (const [key, value] of Object.entries(patch)) {
    if (value !== undefined) {
      (store as unknown as Record<string, unknown>)[key] = value;
    }
  }
}
