/**
 * Rollout control for step-wise Tier 3 execution (TASKS.md #153,
 * STEPWISE_EXECUTION.md).
 *
 * When on, a Tier 3 request plans the whole build up front but executes and
 * previews only ONE dependency wave per HTTP request, gated on the client
 * posting a decision to `/excel-ai/conversation/continue` before the next wave
 * is generated. When off, Tier 3 behaves exactly as it did before — one
 * request, whole build, progressive cards via `onWaveComplete`.
 *
 * Env: `ENABLE_STEPWISE_EXECUTION=on|off`
 * Aliases: true/1 → on; false/0 → off.
 *
 * Defaults **off everywhere, including development**, unlike
 * ENABLE_COMPLEXITY_TIERING. This changes the shape of the SSE contract a
 * client must implement (a run now ends with `wave_ready`, not
 * `conversation_end`), so an add-in build that predates the client half must
 * not silently start receiving paused runs it will never continue.
 */
export type StepwiseExecutionMode = 'on' | 'off';

export function parseStepwiseExecutionMode(
  raw: string | undefined,
): StepwiseExecutionMode {
  const value = (raw ?? '').trim().toLowerCase();
  if (value === 'on' || value === 'true' || value === '1') return 'on';
  return 'off';
}

export function getStepwiseExecutionMode(): StepwiseExecutionMode {
  return parseStepwiseExecutionMode(process.env.ENABLE_STEPWISE_EXECUTION);
}

export function isStepwiseExecutionEnabled(
  mode: StepwiseExecutionMode = getStepwiseExecutionMode(),
): boolean {
  return mode === 'on';
}

/**
 * Stepwise only pays for itself on builds with more than one dependency wave —
 * a single-wave plan gated on an accept is just today's behaviour plus a
 * round trip, so it stays one-shot.
 */
export function shouldRunStepwise(waveCount: number): boolean {
  return isStepwiseExecutionEnabled() && waveCount > 1;
}
