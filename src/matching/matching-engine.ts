/**
 * P2.1 / B6 — generic A-vs-B matching engine (deterministic, no LLM).
 *
 * Building blocks every reconciliation-like feature can share:
 *  - name normalisation for Indian party names (M/s, Pvt/Private, Ltd/Limited, LLP, &/and…)
 *  - a 0–100 similarity score (edit distance + fuzzy token overlap)
 *  - confidence bands (≥95 high · 80–95 review · <80 none, configurable)
 *  - one-to-one best-first assignment (each row used once)
 *  - many-to-one grouping by amount (unique subset-sum, never guesses between ties)
 *  - amount tolerance (absolute ₹ and/or %)
 *
 * The existing GST and bank reconciliation passes are NOT rewired onto this module
 * (their specs pin exact behaviour); new features (P2.2 fuzzy match, P3.9 many-to-one,
 * the colleague's C3/C6/C7 items) build on it.
 */
import { normalizeInvoiceKey } from '../excel-ai/utils/india-validators.util';

export interface Bands {
  high: number;
  review: number;
}
export const DEFAULT_BANDS: Bands = { high: 95, review: 80 };
export type Band = 'high' | 'review' | 'none';

const LEGAL_SUFFIXES = /\b(messrs|pvt|private|ltd|limited|llp|co|company|corp|corporation|inc|plc|opc)\b\.?/gi;

/** "M/s ABC Private Limited." → "abc"; "Kumar & Sons" → "kumar and sons". */
export function normalizePartyName(raw: unknown): string {
  return String(raw ?? '')
    .toLowerCase()
    .replace(/ /g, ' ')
    .replace(/&/g, ' and ')
    .replace(/\bm\/s\.?/g, ' ')
    .replace(LEGAL_SUFFIXES, ' ')
    .replace(/[^a-z0-9 ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i += 1) {
    const cur = [i];
    for (let j = 1; j <= b.length; j += 1) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[b.length];
}

const ratio = (a: string, b: string): number =>
  a.length + b.length === 0 ? 1 : 1 - levenshtein(a, b) / Math.max(a.length, b.length);

/** Tokens count as equal when identical, one is a short stem/plural of the other, or 1 edit apart (len ≥ 5). */
function tokensMatch(x: string, y: string): boolean {
  if (x === y) return true;
  if (x.length >= 4 && y.length >= 4 && (x.startsWith(y) || y.startsWith(x)) && Math.abs(x.length - y.length) <= 2) {
    return true;
  }
  return Math.min(x.length, y.length) >= 5 && levenshtein(x, y) <= 1;
}

/** 0–100. Max of whole-string edit ratio and fuzzy token Dice overlap, on normalised values. */
export function similarity(a: unknown, b: unknown, normalise: (v: unknown) => string = normalizePartyName): number {
  const x = normalise(a);
  const y = normalise(b);
  if (!x || !y) return 0;
  if (x === y) return 100;
  const tx = x.split(' ');
  const ty = y.split(' ');
  const used = new Set<number>();
  let hits = 0;
  for (const t of tx) {
    const j = ty.findIndex((u, k) => !used.has(k) && tokensMatch(t, u));
    if (j >= 0) {
      used.add(j);
      hits += 1;
    }
  }
  const dice = (2 * hits) / (tx.length + ty.length);
  return Math.round(Math.max(ratio(x, y), dice) * 100);
}

export function bandOf(score: number, bands: Bands = DEFAULT_BANDS): Band {
  if (score >= bands.high) return 'high';
  if (score >= bands.review) return 'review';
  return 'none';
}

export interface Candidate<L, R> {
  left: L;
  leftIndex: number;
  right: R | null;
  rightIndex: number | null;
  score: number;
  band: Band;
}

/**
 * One-to-one: every pair scored, best pairs taken first, each side used once.
 * Ties resolved by original order, so the result is deterministic.
 */
export function matchOneToOne<L, R>(
  left: L[],
  right: R[],
  score: (l: L, r: R) => number,
  bands: Bands = DEFAULT_BANDS,
): Candidate<L, R>[] {
  const pairs: Array<{ i: number; j: number; s: number }> = [];
  left.forEach((l, i) => right.forEach((r, j) => pairs.push({ i, j, s: score(l, r) })));
  pairs.sort((p, q) => q.s - p.s || p.i - q.i || p.j - q.j);
  const usedL = new Set<number>();
  const usedR = new Set<number>();
  const best = new Map<number, { j: number; s: number }>();
  for (const p of pairs) {
    if (usedL.has(p.i) || usedR.has(p.j) || bandOf(p.s, bands) === 'none') continue;
    usedL.add(p.i);
    usedR.add(p.j);
    best.set(p.i, { j: p.j, s: p.s });
  }
  return left.map((l, i) => {
    const hit = best.get(i);
    return hit
      ? { left: l, leftIndex: i, right: right[hit.j], rightIndex: hit.j, score: hit.s, band: bandOf(hit.s, bands) }
      : { left: l, leftIndex: i, right: null, rightIndex: null, score: 0, band: 'none' as Band };
  });
}

export interface Tolerance {
  /** Absolute ₹ difference allowed (inclusive). */
  absolute?: number;
  /** Percentage of the larger amount allowed (inclusive). */
  percent?: number;
}

export function withinTolerance(a: number, b: number, tol: Tolerance = {}): boolean {
  const diff = Math.abs(a - b);
  if (diff < 0.005) return true;
  if (tol.absolute !== undefined && diff <= tol.absolute + 1e-9) return true;
  if (tol.percent !== undefined && diff <= (Math.max(Math.abs(a), Math.abs(b)) * tol.percent) / 100 + 1e-9) return true;
  return false;
}

/**
 * Many-to-one: indexes of `amounts` (≥ 2 of them, at most `maxSize`) that sum exactly
 * (to the paisa) to `target` — returned only when the solution is UNIQUE; ambiguity → null.
 */
export function uniqueSubsetSum(amounts: number[], target: number, maxSize = 5, nodeBudget = 150_000): number[] | null {
  const cents = amounts.map((a) => Math.round(a * 100));
  const goal = Math.round(target * 100);
  const order = cents
    .map((_, i) => i)
    .filter((i) => cents[i] > 0)
    .sort((x, y) => cents[y] - cents[x] || x - y);
  const solutions: number[][] = [];
  let nodes = 0;
  const pick: number[] = [];
  const dfs = (start: number, remaining: number): void => {
    if (solutions.length > 1 || nodes > nodeBudget) return;
    nodes += 1;
    if (remaining === 0) {
      if (pick.length >= 2) solutions.push([...pick]);
      return;
    }
    if (pick.length >= maxSize) return;
    for (let k = start; k < order.length; k += 1) {
      const v = cents[order[k]];
      if (v > remaining) continue;
      pick.push(order[k]);
      dfs(k + 1, remaining - v);
      pick.pop();
      if (solutions.length > 1) return;
    }
  };
  dfs(0, goal);
  if (nodes > nodeBudget || solutions.length !== 1) return null;
  return solutions[0].sort((a, b) => a - b);
}

/** Invoice numbers compared on the shared B10 key ("INV/24-25/0233" = "233"). */
export const invoiceSimilarity = (a: unknown, b: unknown): number => {
  const x = normalizeInvoiceKey(a);
  const y = normalizeInvoiceKey(b);
  if (!x || !y) return 0;
  return x === y ? 100 : Math.round(ratio(x, y) * 100);
};
