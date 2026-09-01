/**
 * TASKS.md #74 — performance/latency/token rollups, segmented by route and complexity
 * tier, so `PRD.md` §6.3's "an aggregate number masks the case that matters" applies to
 * speed the same way it already applies to Tier A safety metrics (`tier-a-metrics.util.ts`).
 *
 * Pure computation, no DB access — `performance-metrics.service.ts` turns `workflow_traces`
 * documents and `ContextCacheService.getStats()` into the plain records these functions
 * consume, following the exact split `tier-a-metrics.util.ts`/`.service.ts` already
 * established.
 *
 * Total request latency is DERIVED, not read off a stored field: `workflow_traces` has no
 * trace-level duration, only per-node `startedAt`/`endedAt`/`durationMs`. Using
 * `max(endedAt) - min(startedAt)` across a trace's nodes (rather than summing each node's
 * own `durationMs`) is deliberate — summing would double-count if any nodes ever overlap
 * (e.g. a future parallel-subtask change, TASKS.md #77), while the min/max span is correct
 * regardless of whether nodes ran sequentially or concurrently.
 */

export interface TraceLatencyRecord {
  route?: string;
  tier?: number;
  /** Derived span in ms — null if the trace has no nodes with both timestamps set. */
  totalDurationMs: number | null;
}

export interface PercentileLatency {
  p50: number | null;
  p95: number | null;
  p99: number | null;
  avg: number | null;
  count: number;
}

export interface LatencyReport {
  overall: PercentileLatency;
  /** Keyed by `segmentKey(route, tier)` — same keying convention as tier-a-metrics.util.ts. */
  segments: Record<string, PercentileLatency>;
}

export interface TokenUsageRecord {
  tier?: string;
  promptTokens: number;
  completionTokens: number;
}

export interface TokenReport {
  totalPromptTokens: number;
  totalCompletionTokens: number;
  totalTokens: number;
  avgTokensPerCall: number | null;
  callCount: number;
  byModelTier: Record<string, { calls: number; promptTokens: number; completionTokens: number }>;
}

export interface CacheStatsSnapshot {
  hits: number;
  misses: number;
  stableHits: number;
  stableMisses: number;
}

export interface CacheReport {
  conversationHitRate: number | null;
  stableHitRate: number | null;
  combinedHitRate: number | null;
}

export interface PerformanceReport {
  latency: LatencyReport;
  tokens: TokenReport;
  cache: CacheReport;
}

function percentile(sorted: number[], p: number): number | null {
  if (sorted.length === 0) return null;
  const index = Math.min(sorted.length - 1, Math.ceil((p / 100) * sorted.length) - 1);
  return sorted[Math.max(0, index)];
}

function computePercentiles(durations: number[]): PercentileLatency {
  if (durations.length === 0) {
    return { p50: null, p95: null, p99: null, avg: null, count: 0 };
  }
  const sorted = [...durations].sort((a, b) => a - b);
  const avg = sorted.reduce((sum, d) => sum + d, 0) / sorted.length;
  return {
    p50: percentile(sorted, 50),
    p95: percentile(sorted, 95),
    p99: percentile(sorted, 99),
    avg,
    count: sorted.length,
  };
}

/** Same segment-key convention as tier-a-metrics.util.ts's segmentKey(). */
export function segmentKey(route: string | undefined, tier: number | undefined): string {
  return `${route ?? 'unknown'}/tier${tier ?? 'unknown'}`;
}

export function buildLatencyReport(records: TraceLatencyRecord[]): LatencyReport {
  const validDurations = records
    .map((r) => r.totalDurationMs)
    .filter((d): d is number => d !== null);

  const bySegment = new Map<string, number[]>();
  for (const record of records) {
    if (record.totalDurationMs === null) continue;
    const key = segmentKey(record.route, record.tier);
    const list = bySegment.get(key) ?? [];
    list.push(record.totalDurationMs);
    bySegment.set(key, list);
  }

  const segments: Record<string, PercentileLatency> = {};
  for (const [key, durations] of bySegment.entries()) {
    segments[key] = computePercentiles(durations);
  }

  return {
    overall: computePercentiles(validDurations),
    segments,
  };
}

export function buildTokenReport(records: TokenUsageRecord[]): TokenReport {
  const totalPromptTokens = records.reduce((sum, r) => sum + r.promptTokens, 0);
  const totalCompletionTokens = records.reduce((sum, r) => sum + r.completionTokens, 0);
  const totalTokens = totalPromptTokens + totalCompletionTokens;
  const callCount = records.length;
  const avgTokensPerCall = callCount > 0 ? totalTokens / callCount : null;

  const byModelTier: Record<
    string,
    { calls: number; promptTokens: number; completionTokens: number }
  > = {};
  for (const record of records) {
    const tier = record.tier ?? 'unknown';
    if (!byModelTier[tier]) {
      byModelTier[tier] = { calls: 0, promptTokens: 0, completionTokens: 0 };
    }
    byModelTier[tier].calls += 1;
    byModelTier[tier].promptTokens += record.promptTokens;
    byModelTier[tier].completionTokens += record.completionTokens;
  }

  return {
    totalPromptTokens,
    totalCompletionTokens,
    totalTokens,
    avgTokensPerCall,
    callCount,
    byModelTier,
  };
}

/** Cache stats come from ContextCacheService.getStats() (TASKS.md #68) — a live in-memory
 * process snapshot, not a date-ranged query. Reported as-is; the caller decides how to
 * present "since process start" vs. a specific window (the metric isn't windowable today
 * since the cache tracks lifetime counters, not per-request timestamps). */
export function buildCacheReport(stats: CacheStatsSnapshot): CacheReport {
  const conversationTotal = stats.hits + stats.misses;
  const stableTotal = stats.stableHits + stats.stableMisses;
  const combinedTotal = conversationTotal + stats.stableHits;

  return {
    conversationHitRate: conversationTotal > 0 ? stats.hits / conversationTotal : null,
    stableHitRate: stableTotal > 0 ? stats.stableHits / stableTotal : null,
    // "Combined" = how often ANY cache layer served the request without a full rebuild —
    // a conversation-hit already short-circuits before the stable layer is even checked
    // (see context-cache.service.ts's get()), so this is hits+stableHits over
    // conversationTotal, not a simple sum of two independent rates.
    combinedHitRate: conversationTotal > 0 ? (stats.hits + stats.stableHits) / conversationTotal : null,
  };
}

export function buildPerformanceReport(
  traceRecords: TraceLatencyRecord[],
  tokenRecords: TokenUsageRecord[],
  cacheStats: CacheStatsSnapshot,
): PerformanceReport {
  return {
    latency: buildLatencyReport(traceRecords),
    tokens: buildTokenReport(tokenRecords),
    cache: buildCacheReport(cacheStats),
  };
}
