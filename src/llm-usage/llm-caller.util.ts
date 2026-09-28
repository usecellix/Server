/**
 * Derives "which part of the pipeline made this LLM call" from the call stack,
 * so new call sites are labelled automatically instead of each having to pass a
 * purpose string. Must be called synchronously at the public entry point of
 * OpenRouterService, before any await, while the caller's frame is still on
 * the stack.
 *
 * `planner.agent.ts` → `planner`, `llm-router.service.ts` → `llm-router`.
 */
const SKIP_FILES = ['openrouter.service', 'llm-caller.util', 'node:', 'node_modules'];

export function llmCallerFromStack(stack: string | undefined): string {
  if (!stack) return 'unknown';
  for (const line of stack.split('\n').slice(1)) {
    const match = /[\\/]([\w.-]+)\.(?:ts|js):\d+:\d+\)?\s*$/.exec(line);
    if (!match) continue;
    if (SKIP_FILES.some((skip) => line.includes(skip))) continue;
    return match[1].replace(/\.(service|agent|util|controller)$/, '');
  }
  return 'unknown';
}
