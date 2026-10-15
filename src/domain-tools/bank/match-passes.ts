import {
  BankMatch,
  BankMatchResult,
  BankMatchSettings,
  BankTxn,
  DEFAULT_BANK_MATCH_SETTINGS,
  Direction,
} from '../../bank-recon/types';
import { daysDiff } from '../ingestion/bank/bank-utils';
import { diagnoseUnmatched } from './diagnose-unmatched';
import { findGroupMatches } from './one-to-many';
import { narrationSimilarity, normalizeRef } from './narration-parser';

/**
 * Bank reconciliation matching pipeline (plan §6). Pure and deterministic: the same inputs
 * always yield the same output. Direction is already normalized (IN / OUT) on every row, so
 * a books "Debit" and a bank "Credit" are compared like with like here.
 */

export interface MatchInput {
  books: BankTxn[];
  bank: BankTxn[];
  settings?: Partial<BankMatchSettings>;
  /** ISO period end — drives timing-difference diagnosis. */
  periodEnd?: string;
}

const RETURN_RE = /\b(return(ed)?|revers(al|ed)?|bounce[d]?|unpaid|failed|refund(ed)?|rtn|reject(ed)?|dishon(ou)?r(ed)?)\b/i;

const cents = (n: number) => Math.round(n * 100);
const keyOf = (dir: Direction, amount: number) => `${dir}|${cents(amount)}`;

function refTokens(t: BankTxn): string[] {
  const out: string[] = [];
  const a = normalizeRef(t.instrumentNo);
  const b = normalizeRef(t.referenceNo);
  if (a.length >= 3) out.push(`I:${a}`);
  if (b.length >= 6) out.push(`R:${b}`);
  return out;
}

export function refMatch(a: BankTxn, b: BankTxn): boolean {
  const tb = new Set(refTokens(b));
  return refTokens(a).some((x) => tb.has(x));
}

export function refConflict(a: BankTxn, b: BankTxn): boolean {
  const ai = normalizeRef(a.instrumentNo);
  const bi = normalizeRef(b.instrumentNo);
  if (ai.length >= 3 && bi.length >= 3 && ai !== bi) return true;
  const ar = normalizeRef(a.referenceNo);
  const br = normalizeRef(b.referenceNo);
  return ar.length >= 6 && br.length >= 6 && ar !== br;
}

function similarity(a: BankTxn, b: BankTxn): number {
  let s = narrationSimilarity(a.narration, b.narration);
  if (a.counterparty && b.counterparty) {
    s = Math.max(s, narrationSimilarity(a.counterparty, b.counterparty));
  } else if (a.counterparty) {
    s = Math.max(s, narrationSimilarity(a.counterparty, b.narration) * 0.9);
  } else if (b.counterparty) {
    s = Math.max(s, narrationSimilarity(a.narration, b.counterparty) * 0.9);
  }
  return s;
}

/** Hungarian algorithm (min-cost assignment) on a rectangular cost matrix. */
export function hungarian(cost: number[][]): number[] {
  const n = cost.length;
  const m = cost[0]?.length ?? 0;
  const size = Math.max(n, m);
  const INF = 1e9;
  const a: number[][] = Array.from({ length: size + 1 }, () => new Array(size + 1).fill(0));
  for (let i = 1; i <= size; i++) {
    for (let j = 1; j <= size; j++) {
      a[i][j] = i <= n && j <= m ? cost[i - 1][j - 1] : INF;
    }
  }
  const u = new Array(size + 1).fill(0);
  const v = new Array(size + 1).fill(0);
  const p = new Array(size + 1).fill(0);
  const way = new Array(size + 1).fill(0);
  for (let i = 1; i <= size; i++) {
    p[0] = i;
    let j0 = 0;
    const minv = new Array(size + 1).fill(Infinity);
    const used = new Array(size + 1).fill(false);
    do {
      used[j0] = true;
      const i0 = p[j0];
      let delta = Infinity;
      let j1 = 0;
      for (let j = 1; j <= size; j++) {
        if (used[j]) continue;
        const cur = a[i0][j] - u[i0] - v[j];
        if (cur < minv[j]) {
          minv[j] = cur;
          way[j] = j0;
        }
        if (minv[j] < delta) {
          delta = minv[j];
          j1 = j;
        }
      }
      for (let j = 0; j <= size; j++) {
        if (used[j]) {
          u[p[j]] += delta;
          v[j] -= delta;
        } else minv[j] -= delta;
      }
      j0 = j1;
    } while (p[j0] !== 0);
    do {
      const j1 = way[j0];
      p[j0] = p[j1];
      j0 = j1;
    } while (j0);
  }
  const assign = new Array(n).fill(-1);
  for (let j = 1; j <= size; j++) {
    if (p[j] >= 1 && p[j] <= n && j <= m) assign[p[j] - 1] = j - 1;
  }
  return assign;
}

export function bankMatch(input: MatchInput): BankMatchResult {
  const settings: BankMatchSettings = { ...DEFAULT_BANK_MATCH_SETTINGS, ...input.settings };
  const { books, bank } = input;
  const matches: BankMatch[] = [];
  const usedBooks = new Set<string>();
  const usedBank = new Set<string>();
  let counter = 0;

  const tol = Math.round(settings.amountTolerance * 100);
  const amountOk = (a: BankTxn, b: BankTxn) => Math.abs(cents(a.amount) - cents(b.amount)) <= tol;
  // bank − books, in days
  const gapOf = (bk: BankTxn, bn: BankTxn) => daysDiff(bk.txnDate, bn.txnDate);
  const inClearing = (g: number) => g >= -settings.earlyWindowDays && g <= settings.clearingWindowDays;
  const inTight = (g: number) => Math.abs(g) <= settings.tightWindowDays;

  const bankIndex = new Map<string, BankTxn[]>();
  for (const t of bank) {
    const k = keyOf(t.direction, t.amount);
    (bankIndex.get(k) ?? bankIndex.set(k, []).get(k)!).push(t);
  }
  // With a non-zero tolerance exact-key lookup is not enough — scan the direction bucket.
  const bankByDir: Record<Direction, BankTxn[]> = { IN: [], OUT: [] };
  for (const t of bank) bankByDir[t.direction].push(t);

  const candidatesFor = (b: BankTxn): BankTxn[] => {
    const pool = tol === 0 ? (bankIndex.get(keyOf(b.direction, b.amount)) ?? []) : bankByDir[b.direction];
    return pool.filter((x) => !usedBank.has(x.id) && amountOk(b, x));
  };

  const link = (
    pass: number,
    confidence: number,
    b: BankTxn,
    k: BankTxn,
  ) => {
    usedBooks.add(b.id);
    usedBank.add(k.id);
    matches.push({ id: `m${++counter}`, kind: 'ONE_TO_ONE', pass, confidence, books: [b], bank: [k] });
  };

  // ---- Pass 1: exact reference ----------------------------------------------------------
  {
    const byToken = new Map<string, BankTxn[]>();
    for (const k of bank) for (const tk of refTokens(k)) (byToken.get(tk) ?? byToken.set(tk, []).get(tk)!).push(k);
    const pairs: Array<{ b: BankTxn; k: BankTxn; cost: number }> = [];
    for (const b of books) {
      const seen = new Set<string>();
      for (const tk of refTokens(b)) {
        for (const k of byToken.get(tk) ?? []) {
          if (seen.has(k.id)) continue;
          seen.add(k.id);
          if (k.direction !== b.direction || !amountOk(b, k)) continue;
          const g = gapOf(b, k);
          if (!inClearing(g)) continue;
          pairs.push({ b, k, cost: Math.abs(g) });
        }
      }
    }
    pairs.sort((x, y) => x.cost - y.cost);
    for (const p of pairs) {
      if (usedBooks.has(p.b.id) || usedBank.has(p.k.id)) continue;
      link(1, 1, p.b, p.k);
    }
  }

  // ---- Pass 2: amount + date, only when unambiguous in BOTH directions ----------------------
  {
    const candB = new Map<string, BankTxn[]>();
    const candK = new Map<string, BankTxn[]>();
    for (const b of books) {
      if (usedBooks.has(b.id)) continue;
      const cs = candidatesFor(b).filter((k) => inTight(gapOf(b, k)) && !refConflict(b, k));
      candB.set(b.id, cs);
      for (const k of cs) (candK.get(k.id) ?? candK.set(k.id, []).get(k.id)!).push(b);
    }
    for (const b of books) {
      if (usedBooks.has(b.id)) continue;
      const cs = candB.get(b.id) ?? [];
      if (cs.length !== 1) continue; // ties go to Pass 3/4 — never guess
      const k = cs[0];
      if (usedBank.has(k.id) || (candK.get(k.id) ?? []).length !== 1) continue;
      link(2, 0.9, b, k);
    }
  }

  // ---- Pass 3: narration-assisted ---------------------------------------------------------
  {
    const pairs: Array<{ b: BankTxn; k: BankTxn; sim: number; gap: number }> = [];
    for (const b of books) {
      if (usedBooks.has(b.id)) continue;
      for (const k of candidatesFor(b)) {
        const g = gapOf(b, k);
        if (!inClearing(g) || refConflict(b, k)) continue;
        const sim = similarity(b, k);
        if (sim >= settings.narrationThreshold) pairs.push({ b, k, sim, gap: Math.abs(g) });
      }
    }
    pairs.sort((x, y) => y.sim - x.sim || x.gap - y.gap);
    for (const p of pairs) {
      if (usedBooks.has(p.b.id) || usedBank.has(p.k.id)) continue;
      link(3, 0.85, p.b, p.k);
    }
  }

  // ---- Pass 4: optimal one-to-one assignment within identical (direction, amount) groups ----
  {
    const groups = new Map<string, { b: BankTxn[]; k: BankTxn[] }>();
    for (const b of books) {
      if (usedBooks.has(b.id)) continue;
      const g = groups.get(keyOf(b.direction, b.amount)) ?? { b: [], k: [] };
      g.b.push(b);
      groups.set(keyOf(b.direction, b.amount), g);
    }
    for (const [key, g] of groups) {
      g.k = (bankIndex.get(key) ?? []).filter((x) => !usedBank.has(x.id));
      if (!g.k.length || tol !== 0) continue;
      const BIG = 1e6;
      const cost = g.b.map((b) =>
        g.k.map((k) => {
          const gap = gapOf(b, k);
          if (!inClearing(gap) || refConflict(b, k)) return BIG;
          return Math.abs(gap) + (1 - similarity(b, k)) * 2;
        }),
      );
      const assign = hungarian(cost);
      assign.forEach((j, i) => {
        if (j < 0 || cost[i][j] >= BIG) return;
        link(4, 0.6, g.b[i], g.k[j]);
      });
    }
  }

  // ---- Pass 5: one-to-many / many-to-one ---------------------------------------------------
  if (settings.enableGroupMatch) {
    const found = findGroupMatches({
      books: books.filter((b) => !usedBooks.has(b.id)),
      bank: bank.filter((k) => !usedBank.has(k.id)),
      maxGroupSize: settings.maxGroupSize,
      inWindow: (b, k) => inClearing(gapOf(b, k)),
    });
    for (const g of found) {
      g.books.forEach((b) => usedBooks.add(b.id));
      g.bank.forEach((k) => usedBank.add(k.id));
      matches.push({ id: `m${++counter}`, kind: 'GROUP', pass: 5, confidence: 0.7, books: g.books, bank: g.bank });
    }
  }

  // ---- Pass 7: reversal / bounce pairing within a single side --------------------------------
  const pairReversals = (rows: BankTxn[], used: Set<string>, side: 'BOOKS' | 'BANK') => {
    const free = rows.filter((r) => !used.has(r.id));
    for (const ret of free) {
      if (used.has(ret.id) || !RETURN_RE.test(ret.narration)) continue;
      let best: BankTxn | undefined;
      let bestGap = Infinity;
      for (const orig of free) {
        if (orig.id === ret.id || used.has(orig.id)) continue;
        if (orig.direction === ret.direction || cents(orig.amount) !== cents(ret.amount)) continue;
        const g = daysDiff(orig.txnDate, ret.txnDate);
        if (g < 0 || g > 30) continue;
        if (g < bestGap) {
          best = orig;
          bestGap = g;
        }
      }
      if (best) {
        used.add(ret.id);
        used.add(best.id);
        matches.push({
          id: `m${++counter}`,
          kind: 'REVERSAL',
          pass: 7,
          confidence: 0.8,
          side,
          books: side === 'BOOKS' ? [best, ret] : [],
          bank: side === 'BANK' ? [best, ret] : [],
        });
      }
    }
  };
  pairReversals(books, usedBooks, 'BOOKS');
  pairReversals(bank, usedBank, 'BANK');

  // ---- Finalization: cross-pass consistency + diagnosis --------------------------------------
  const seen = new Set<string>();
  for (const m of matches) {
    for (const t of [...m.books, ...m.bank]) {
      if (seen.has(t.id)) throw new Error(`bank-recon invariant violated: row ${t.id} used twice`);
      seen.add(t.id);
    }
  }

  const { unmatchedBooks, unmatchedBank } = diagnoseUnmatched({
    allBooks: books,
    allBank: bank,
    freeBooks: books.filter((b) => !usedBooks.has(b.id)),
    freeBank: bank.filter((k) => !usedBank.has(k.id)),
    usedBooks,
    usedBank,
    settings,
    periodEnd: input.periodEnd,
    refMatch,
    refConflict,
    gapOf,
  });

  return { matches, unmatchedBooks, unmatchedBank };
}
