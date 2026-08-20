/**
 * PRD Tier A metrics rollup (TASKS.md #50/#51). Pure computation, no DB access — the
 * service layer (`tier-a-metrics.service.ts`) is responsible for turning `workflow_traces`
 * and `change_sets` documents into the plain records these functions consume.
 *
 * Two scoping decisions, made explicit here rather than left implicit in the numbers:
 *
 * 1. **A1/A4 are `workflow_traces`-derived, not `request_logs`-derived**, despite PRD.md
 *    §6.1 citing `request_logs` as the source. Investigated first: `request_logs` carries
 *    no `route`/`tier`/verification-outcome fields at all (just method/url/statusCode) —
 *    `workflow_traces` is the collection that actually persists route, tier, and a
 *    structured per-node verifier outcome. This is a source correction, not a metric
 *    redefinition — the PRD's *definitions* of A1/A4 are unchanged, only where the data
 *    comes from.
 * 2. **"Verification was skipped" is inferred as the absence of a `verifier`-typed node**,
 *    not read off an explicit boolean — no code path in this repo ever appends a verifier
 *    node with `status: 'skipped'`; a request that never ran verification (Tier 0/1, or a
 *    non-write route) simply has no such node.
 * 3. **A4's PRD-specified "excludes cancelled-at-preview" carve-out is not currently
 *    measurable.** Neither `workflow_traces.status` nor `ChangeSet.status` distinguishes
 *    "user rejected/cancelled before ever accepting" from "user reverted after applying" —
 *    both collapse to `'rejected'` / `'reverted'`. A4 here is computed over write-route
 *    requests that reached *either* terminal decision (`'accepted'` or `'rejected'`),
 *    which folds cancel-at-preview into the same denominator bucket as revert-after-apply.
 *    This is a documented approximation, not a silent one — flagged again in TASKS.md #50.
 */

export interface TraceOutcomeRecord {
  route?: string;
  tier?: number;
  /** `WorkflowTraceStatus` — kept as `string` here so this module has no schema import. */
  status: string;
  verifierNodeStatus: 'success' | 'failed' | 'skipped';
}

export interface ChangeSetOutcomeRecord {
  route?: string;
  tier?: number;
  unintendedCount: number;
  formulaErrorCount: number;
}

export interface MetricResult {
  /** null when the denominator is zero — "not enough data," not "0%." */
  rate: number | null;
  numerator: number;
  denominator: number;
}

export interface TierAMetrics {
  a1FalseSuccessRate: MetricResult;
  a4VerifiedCompletionRate: MetricResult;
  a5UnintendedModificationRate: MetricResult;
  a6FormulaErrorRate: MetricResult;
}

export interface TierAMetricsReport {
  overall: TierAMetrics;
  /** Keyed by `segmentKey(route, tier)` — PRD §6.3's "report per-tier or the number misleads." */
  segments: Record<string, TierAMetrics>;
}

function toRate(numerator: number, denominator: number): MetricResult {
  return { rate: denominator > 0 ? numerator / denominator : null, numerator, denominator };
}

export function segmentKey(route: string | undefined, tier: number | undefined): string {
  return `${route ?? 'unknown'}/tier${tier ?? 'unknown'}`;
}

/** A1 — requests reported complete (write-route, reached 'accepted') where verification
 * failed or never ran, over all write-route requests reported complete. */
export function computeA1FalseSuccessRate(records: TraceOutcomeRecord[]): MetricResult {
  const reportedComplete = records.filter((r) => r.route === 'write' && r.status === 'accepted');
  const falseSuccess = reportedComplete.filter((r) => r.verifierNodeStatus !== 'success');
  return toRate(falseSuccess.length, reportedComplete.length);
}

/** A4 — write-route requests that reached a terminal decision (accepted or rejected;
 * see module doc point 3 on why "rejected" also stands in for "cancelled at preview")
 * where the outcome was a verified, applied success. */
export function computeA4VerifiedCompletionRate(records: TraceOutcomeRecord[]): MetricResult {
  const terminal = records.filter(
    (r) => r.route === 'write' && (r.status === 'accepted' || r.status === 'rejected'),
  );
  const verifiedComplete = terminal.filter(
    (r) => r.status === 'accepted' && r.verifierNodeStatus === 'success',
  );
  return toRate(verifiedComplete.length, terminal.length);
}

/** A5 — applied change sets containing >=1 unintended (out-of-scope) cell change. */
export function computeA5UnintendedModificationRate(records: ChangeSetOutcomeRecord[]): MetricResult {
  return toRate(records.filter((r) => r.unintendedCount > 0).length, records.length);
}

/** A6 — applied change sets that introduced >=1 new Excel-error cell. */
export function computeA6FormulaErrorRate(records: ChangeSetOutcomeRecord[]): MetricResult {
  return toRate(records.filter((r) => r.formulaErrorCount > 0).length, records.length);
}

function computeMetricsFor(
  traces: TraceOutcomeRecord[],
  changeSets: ChangeSetOutcomeRecord[],
): TierAMetrics {
  return {
    a1FalseSuccessRate: computeA1FalseSuccessRate(traces),
    a4VerifiedCompletionRate: computeA4VerifiedCompletionRate(traces),
    a5UnintendedModificationRate: computeA5UnintendedModificationRate(changeSets),
    a6FormulaErrorRate: computeA6FormulaErrorRate(changeSets),
  };
}

/**
 * TASKS.md #51 — segment by route and complexity tier, not just one blended number
 * (PRD §6.3). A segment is included whenever either record type has at least one entry
 * with that (route, tier) pair — trace-derived (A1/A4) and change-set-derived (A5/A6)
 * segments are independent sets that happen to share a keying scheme, not a join.
 */
export function buildTierAMetricsReport(
  traceRecords: TraceOutcomeRecord[],
  changeSetRecords: ChangeSetOutcomeRecord[],
): TierAMetricsReport {
  const overall = computeMetricsFor(traceRecords, changeSetRecords);

  const keys = new Set<string>();
  for (const r of traceRecords) keys.add(segmentKey(r.route, r.tier));
  for (const r of changeSetRecords) keys.add(segmentKey(r.route, r.tier));

  const segments: Record<string, TierAMetrics> = {};
  for (const key of keys) {
    segments[key] = computeMetricsFor(
      traceRecords.filter((r) => segmentKey(r.route, r.tier) === key),
      changeSetRecords.filter((r) => segmentKey(r.route, r.tier) === key),
    );
  }

  return { overall, segments };
}
