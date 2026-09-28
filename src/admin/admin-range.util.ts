/** Mirrors Dashboard/src/lib/range.ts's RANGES/resolveRange exactly — same keys, same bucket rule. */
const RANGE_MS: Record<string, number | null> = {
  '24h': 24 * 3600_000,
  '7d': 7 * 86_400_000,
  '30d': 30 * 86_400_000,
  '90d': 90 * 86_400_000,
  all: null,
};

export interface ResolvedAdminRange {
  key: string;
  from: Date | null;
  to: Date;
  bucket: 'hour' | 'day';
}

export function resolveAdminRange(rangeKey: string | undefined, now = new Date()): ResolvedAdminRange {
  const key = rangeKey && rangeKey in RANGE_MS ? rangeKey : '30d';
  const ms = RANGE_MS[key];
  return {
    key,
    from: ms === null ? null : new Date(now.getTime() - ms),
    to: now,
    bucket: key === '24h' ? 'hour' : 'day',
  };
}
