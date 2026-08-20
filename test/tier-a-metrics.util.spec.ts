import {
  buildTierAMetricsReport,
  ChangeSetOutcomeRecord,
  computeA1FalseSuccessRate,
  computeA4VerifiedCompletionRate,
  computeA5UnintendedModificationRate,
  computeA6FormulaErrorRate,
  segmentKey,
  TraceOutcomeRecord,
} from '../src/audit/tier-a-metrics.util';

describe('tier-a-metrics.util (TASKS.md #48-51 — PRD Tier A)', () => {
  describe('computeA1FalseSuccessRate', () => {
    it('counts a write-route request reported "accepted" with a failed verifier as false success', () => {
      const records: TraceOutcomeRecord[] = [
        { route: 'write', status: 'accepted', verifierNodeStatus: 'success' },
        { route: 'write', status: 'accepted', verifierNodeStatus: 'failed' },
        { route: 'write', status: 'accepted', verifierNodeStatus: 'skipped' },
      ];
      const result = computeA1FalseSuccessRate(records);
      expect(result).toEqual({ rate: 2 / 3, numerator: 2, denominator: 3 });
    });

    it('ignores non-write routes and non-accepted requests entirely', () => {
      const records: TraceOutcomeRecord[] = [
        { route: 'ask', status: 'accepted', verifierNodeStatus: 'skipped' },
        { route: 'write', status: 'running', verifierNodeStatus: 'skipped' },
        { route: 'write', status: 'rejected', verifierNodeStatus: 'skipped' },
      ];
      expect(computeA1FalseSuccessRate(records)).toEqual({ rate: null, numerator: 0, denominator: 0 });
    });

    it('is 0% (not null) for an all-verified batch', () => {
      const records: TraceOutcomeRecord[] = [
        { route: 'write', status: 'accepted', verifierNodeStatus: 'success' },
      ];
      expect(computeA1FalseSuccessRate(records)).toEqual({ rate: 0, numerator: 0, denominator: 1 });
    });
  });

  describe('computeA4VerifiedCompletionRate', () => {
    it('counts accepted+verified over every write-route request that reached a terminal decision', () => {
      const records: TraceOutcomeRecord[] = [
        { route: 'write', status: 'accepted', verifierNodeStatus: 'success' },
        { route: 'write', status: 'accepted', verifierNodeStatus: 'failed' },
        { route: 'write', status: 'rejected', verifierNodeStatus: 'skipped' },
        { route: 'write', status: 'running', verifierNodeStatus: 'skipped' }, // in-flight, excluded
      ];
      const result = computeA4VerifiedCompletionRate(records);
      expect(result).toEqual({ rate: 1 / 3, numerator: 1, denominator: 3 });
    });

    it('excludes non-write routes', () => {
      const records: TraceOutcomeRecord[] = [
        { route: 'ask', status: 'accepted', verifierNodeStatus: 'success' },
      ];
      expect(computeA4VerifiedCompletionRate(records)).toEqual({ rate: null, numerator: 0, denominator: 0 });
    });
  });

  describe('computeA5UnintendedModificationRate', () => {
    it('flags applied change sets with >=1 unintended change', () => {
      const records: ChangeSetOutcomeRecord[] = [
        { unintendedCount: 0, formulaErrorCount: 0 },
        { unintendedCount: 2, formulaErrorCount: 0 },
        { unintendedCount: 0, formulaErrorCount: 0 },
      ];
      expect(computeA5UnintendedModificationRate(records)).toEqual({
        rate: 1 / 3,
        numerator: 1,
        denominator: 3,
      });
    });

    it('reports null (not 0) when there are zero applied change sets in range', () => {
      expect(computeA5UnintendedModificationRate([])).toEqual({ rate: null, numerator: 0, denominator: 0 });
    });
  });

  describe('computeA6FormulaErrorRate', () => {
    it('flags applied change sets that introduced >=1 new formula error', () => {
      const records: ChangeSetOutcomeRecord[] = [
        { unintendedCount: 0, formulaErrorCount: 1 },
        { unintendedCount: 0, formulaErrorCount: 0 },
      ];
      expect(computeA6FormulaErrorRate(records)).toEqual({ rate: 0.5, numerator: 1, denominator: 2 });
    });
  });

  describe('segmentKey', () => {
    it('falls back to "unknown" for a missing route or tier', () => {
      expect(segmentKey('write', 2)).toBe('write/tier2');
      expect(segmentKey(undefined, 2)).toBe('unknown/tier2');
      expect(segmentKey('write', undefined)).toBe('write/tierunknown');
    });
  });

  describe('buildTierAMetricsReport (TASKS.md #51 — segmentation)', () => {
    it('reports a blended overall number that a per-segment breakdown reveals is misleading (PRD §6.3)', () => {
      // Tier 1 is perfect; Tier 3 is failing badly. The overall number should look
      // fine while masking that — the exact failure mode PRD §6.3 calls out.
      const traceRecords: TraceOutcomeRecord[] = [
        { route: 'write', tier: 1, status: 'accepted', verifierNodeStatus: 'success' },
        { route: 'write', tier: 1, status: 'accepted', verifierNodeStatus: 'success' },
        { route: 'write', tier: 1, status: 'accepted', verifierNodeStatus: 'success' },
        { route: 'write', tier: 1, status: 'accepted', verifierNodeStatus: 'success' },
        { route: 'write', tier: 3, status: 'accepted', verifierNodeStatus: 'failed' },
      ];
      const report = buildTierAMetricsReport(traceRecords, []);

      expect(report.overall.a1FalseSuccessRate.rate).toBe(0.2); // looks tolerable blended
      expect(report.segments['write/tier1'].a1FalseSuccessRate.rate).toBe(0); // actually fine
      expect(report.segments['write/tier3'].a1FalseSuccessRate.rate).toBe(1); // actually broken
    });

    it('keys change-set-derived and trace-derived segments independently by (route, tier)', () => {
      const traceRecords: TraceOutcomeRecord[] = [
        { route: 'write', tier: 2, status: 'accepted', verifierNodeStatus: 'success' },
      ];
      const changeSetRecords: ChangeSetOutcomeRecord[] = [
        { route: 'write', tier: 2, unintendedCount: 1, formulaErrorCount: 0 },
      ];
      const report = buildTierAMetricsReport(traceRecords, changeSetRecords);

      expect(Object.keys(report.segments)).toEqual(['write/tier2']);
      expect(report.segments['write/tier2'].a1FalseSuccessRate).toEqual({
        rate: 0,
        numerator: 0,
        denominator: 1,
      });
      expect(report.segments['write/tier2'].a5UnintendedModificationRate).toEqual({
        rate: 1,
        numerator: 1,
        denominator: 1,
      });
    });
  });
});
