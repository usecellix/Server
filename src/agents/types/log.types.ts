export interface AgentLogEvent {
  correlationId: string;
  agent: 'planner' | 'executor' | 'verifier' | 'workbook';
  model: string;
  durationMs: number;
  success: boolean;
  tokenUsage?: number;
  rawResponse?: string;
  parsedResponse?: unknown;
  error?: string;
}

export interface TierDecisionLog {
  traceId: string;
  message: string;
  /** Tier that actually executed (after feature-flag gating) */
  tier: 0 | 1 | 2 | 3;
  /** Classifier output before feature-flag gating (shadow / rollout) */
  classifiedTier?: 0 | 1 | 2 | 3;
  tieringMode?: 'off' | 'shadow' | 'tier01' | 'full';
  /** True when classified ≠ executed because of shadow/off/tier01 gating */
  shadowed?: boolean;
  matchedBy: 'regex' | 'llm-fallback';
  actionHint: string;
  llmCallCount: number;
  durationMs: number;
  /**
   * Set when a fast lane produced work too big for itself and handed off to the
   * planner (TASKS.md #165). Records the lane that BAILED, so telemetry can
   * separate "the classifier guessed low" from "tier 3 was chosen up front" —
   * the two look identical in `tier` alone, since both end at 3.
   */
  escalatedFrom?: 1 | 2;
  escalationReason?: string;
  /** Actions the run finally produced. Pairs with `tier` to expose mis-sorting. */
  finalActionCount?: number;
}

/** Domain-tool call logged in the same working-paper trail as Planner/Executor/Verifier. */
export interface DomainToolLog {
  traceId: string;
  toolName: string;
  confidence: number;
  exceptionCount: number;
  exceptionCodes: string[];
  sourceRefCount: number;
  durationMs: number;
  success: boolean;
  error?: string;
}
