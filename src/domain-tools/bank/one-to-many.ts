import { BankTxn } from '../../bank-recon/types';
import { narrationSimilarity } from './narration-parser';

/**
 * Pass 5 — group matching: one bank row settling several ledger entries (bulk vendor
 * payment, batched deposit) or the reverse. Bounded subset-sum on exact cents; a group is
 * only reported when the solution is UNIQUE at its size, so the matcher never guesses
 * between equally plausible combinations.
 */

export interface GroupMatchInput {
  books: BankTxn[];
  bank: BankTxn[];
  maxGroupSize: number;
  inWindow: (books: BankTxn, bank: BankTxn) => boolean;
}

export interface GroupMatchResult {
  books: BankTxn[];
  bank: BankTxn[];
}

const MAX_CANDIDATES = 24;
const NODE_BUDGET = 150_000;

const cents = (n: number) => Math.round(n * 100);

function counterpartyCompatible(a: BankTxn, b: BankTxn): boolean {
  if (!a.counterparty || !b.counterparty) return true;
  return narrationSimilarity(a.counterparty, b.counterparty) >= 0.34;
}

/** Returns the unique subset (by index) summing to `target`, or null if none / ambiguous. */
function uniqueSubset(amounts: number[], target: number, size: number): number[] | null {
  const order = amounts.map((_, i) => i).sort((x, y) => amounts[y] - amounts[x]);
  const suffix: number[] = new Array(order.length + 1).fill(0);
  for (let i = order.length - 1; i >= 0; i--) suffix[i] = suffix[i + 1] + amounts[order[i]];

  const solutions: number[][] = [];
  let nodes = 0;
  const pick: number[] = [];

  const dfs = (start: number, remaining: number, left: number) => {
    if (solutions.length > 1 || nodes > NODE_BUDGET) return;
    nodes += 1;
    if (left === 0) {
      if (remaining === 0) solutions.push([...pick]);
      return;
    }
    for (let i = start; i <= order.length - left; i++) {
      const a = amounts[order[i]];
      if (a > remaining) continue;
      if (suffix[i] < remaining) return; // even taking everything left can't reach the target
      pick.push(order[i]);
      dfs(i + 1, remaining - a, left - 1);
      pick.pop();
      if (solutions.length > 1 || nodes > NODE_BUDGET) return;
    }
  };
  dfs(0, target, size);
  if (nodes > NODE_BUDGET) return null;
  return solutions.length === 1 ? solutions[0] : null;
}

export function findGroupMatches(input: GroupMatchInput): GroupMatchResult[] {
  const results: GroupMatchResult[] = [];
  const usedBooks = new Set<string>();
  const usedBank = new Set<string>();

  const run = (anchorsAreBank: boolean) => {
    const anchors = (anchorsAreBank ? input.bank : input.books)
      .filter((a) => !(anchorsAreBank ? usedBank : usedBooks).has(a.id))
      .sort((x, y) => y.amount - x.amount);

    for (const anchor of anchors) {
      if ((anchorsAreBank ? usedBank : usedBooks).has(anchor.id)) continue;
      const pool = (anchorsAreBank ? input.books : input.bank)
        .filter(
          (p) =>
            !(anchorsAreBank ? usedBooks : usedBank).has(p.id) &&
            p.direction === anchor.direction &&
            cents(p.amount) < cents(anchor.amount) &&
            (anchorsAreBank ? input.inWindow(p, anchor) : input.inWindow(anchor, p)) &&
            counterpartyCompatible(anchor, p),
        )
        .sort(
          (x, y) =>
            Math.abs(Date.parse(x.txnDate) - Date.parse(anchor.txnDate)) -
            Math.abs(Date.parse(y.txnDate) - Date.parse(anchor.txnDate)),
        )
        .slice(0, MAX_CANDIDATES);
      if (pool.length < 2) continue;

      const amounts = pool.map((p) => cents(p.amount));
      for (let size = 2; size <= Math.min(input.maxGroupSize, pool.length); size++) {
        const hit = uniqueSubset(amounts, cents(anchor.amount), size);
        if (!hit) continue;
        const members = hit.map((i) => pool[i]);
        members.forEach((m) => (anchorsAreBank ? usedBooks : usedBank).add(m.id));
        (anchorsAreBank ? usedBank : usedBooks).add(anchor.id);
        results.push(
          anchorsAreBank ? { books: members, bank: [anchor] } : { books: [anchor], bank: members },
        );
        break;
      }
    }
  };

  run(true); // one bank row ← many ledger entries
  run(false); // one ledger entry ← many bank rows
  return results;
}
