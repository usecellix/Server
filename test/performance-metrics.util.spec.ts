import {
  buildCacheReport,
  buildLatencyReport,
  buildPerformanceReport,
  buildTokenReport,
  segmentKey,
  TokenUsageRecord,
  TraceLatencyRecord,
} from '../src/audit/performance-metrics.util';

describe('performance-metrics.util (TASKS.md #74)', () => {
  describe('buildLatencyReport', () => {
    it('computes p50/p95/p99/avg over the overall population', () => {
      // 10 durations, evenly spaced 100..1000ms — easy to hand-verify percentiles.
      const records: TraceLatencyRecord[] = Array.from({ length: 10 }, (_, i) => ({
        route: 'write',
        tier: 1,
        totalDurationMs: (i + 1) * 100,
      }));
      const report = buildLatencyReport(records);
      expect(report.overall.count).toBe(10);
      expect(report.overall.avg).toBe(550);
      // p50 of 10 ascending values at index ceil(0.5*10)-1 = 4 → 500
      expect(report.overall.p50).toBe(500);
      // p95 at index ceil(0.95*10)-1 = 9 → 1000
      expect(report.overall.p95).toBe(1000);
    });

    it('excludes null-duration traces (incomplete instrumentation) from percentiles rather than counting them as 0ms', () => {
      const records: TraceLatencyRecord[] = [
        { route: 'write', tier: 1, totalDurationMs: 500 },
        { route: 'write', tier: 1, totalDurationMs: null },
      ];
      const report = buildLatencyReport(records);
      expect(report.overall.count).toBe(1);
      expect(report.overall.avg).toBe(500);
    });

    it('segments by route+tier using the same segmentKey convention as tier-a-metrics.util.ts', () => {
      const records: TraceLatencyRecord[] = [
        { route: 'write', tier: 1, totalDurationMs: 400 },
        { route: 'write', tier: 1, totalDurationMs: 600 },
        { route: 'write', tier: 3, totalDurationMs: 5000 },
      ];
      const report = buildLatencyReport(records);
      expect(Object.keys(report.segments).sort()).toEqual(['write/tier1', 'write/tier3']);
      expect(report.segments['write/tier1'].count).toBe(2);
      expect(report.segments['write/tier1'].avg).toBe(500);
      expect(report.segments['write/tier3'].count).toBe(1);
      expect(report.segments['write/tier3'].avg).toBe(5000);
    });

    it('reports null percentiles (not 0) for an empty population — "no data" must never look like "instant"', () => {
      const report = buildLatencyReport([]);
      expect(report.overall).toEqual({ p50: null, p95: null, p99: null, avg: null, count: 0 });
      expect(report.segments).toEqual({});
    });

    it('segmentKey matches unknown route/tier the same way tier-a-metrics.util.ts does', () => {
      expect(segmentKey(undefined, undefined)).toBe('unknown/tierunknown');
      expect(segmentKey('write', 2)).toBe('write/tier2');
    });
  });

  describe('buildTokenReport', () => {
    it('sums prompt/completion tokens and computes an overall average per call', () => {
      const records: TokenUsageRecord[] = [
        { tier: 'high', promptTokens: 1000, completionTokens: 500 },
        { tier: 'low', promptTokens: 200, completionTokens: 50 },
      ];
      const report = buildTokenReport(records);
      expect(report.totalPromptTokens).toBe(1200);
      expect(report.totalCompletionTokens).toBe(550);
      expect(report.totalTokens).toBe(1750);
      expect(report.callCount).toBe(2);
      expect(report.avgTokensPerCall).toBe(875);
    });

    it('breaks down by LLM model tier (low/medium/high), distinct from the write-route complexity tier', () => {
      const records: TokenUsageRecord[] = [
        { tier: 'high', promptTokens: 1000, completionTokens: 500 },
        { tier: 'high', promptTokens: 800, completionTokens: 400 },
        { tier: 'low', promptTokens: 100, completionTokens: 20 },
      ];
      const report = buildTokenReport(records);
      expect(report.byModelTier.high).toEqual({ calls: 2, promptTokens: 1800, completionTokens: 900 });
      expect(report.byModelTier.low).toEqual({ calls: 1, promptTokens: 100, completionTokens: 20 });
    });

    it('is null (not 0) avgTokensPerCall for zero calls', () => {
      const report = buildTokenReport([]);
      expect(report.avgTokensPerCall).toBeNull();
      expect(report.callCount).toBe(0);
    });
  });

  describe('buildCacheReport', () => {
    it('computes conversation and stable hit rates independently', () => {
      const report = buildCacheReport({ hits: 6, misses: 4, stableHits: 2, stableMisses: 3 });
      expect(report.conversationHitRate).toBe(0.6);
      // stableHitRate is only computed over requests that fell through to the
      // stable-cache check in the first place (conversation-cache misses),
      // matching context-cache.service.ts's own get() short-circuit order.
      expect(report.stableHitRate).toBe(2 / 5);
    });

    it('combinedHitRate reflects "any cache layer served this without a full rebuild"', () => {
      const report = buildCacheReport({ hits: 6, misses: 4, stableHits: 2, stableMisses: 3 });
      // 6 conversation hits + 2 stable hits (which only fire on conversation
      // misses) over the 10 total conversation-level requests = 0.8.
      expect(report.combinedHitRate).toBe(0.8);
    });

    it('returns null rates (not 0) when the cache has served zero requests yet', () => {
      const report = buildCacheReport({ hits: 0, misses: 0, stableHits: 0, stableMisses: 0 });
      expect(report.conversationHitRate).toBeNull();
      expect(report.stableHitRate).toBeNull();
      expect(report.combinedHitRate).toBeNull();
    });
  });

  describe('buildPerformanceReport', () => {
    it('combines latency, token, and cache reports into one payload', () => {
      const report = buildPerformanceReport(
        [{ route: 'write', tier: 1, totalDurationMs: 500 }],
        [{ tier: 'medium', promptTokens: 100, completionTokens: 50 }],
        { hits: 1, misses: 1, stableHits: 0, stableMisses: 1 },
      );
      expect(report.latency.overall.count).toBe(1);
      expect(report.tokens.totalTokens).toBe(150);
      expect(report.cache.conversationHitRate).toBe(0.5);
    });
  });
});
