import { BankMatchResult, BrsResult } from './types';
import { BankReconLayoutDto } from './bank-recon.dto';

/**
 * Holds a completed run's matched result so "change the layout" can re-map it without
 * re-parsing or re-matching. Same short-lived, in-memory, single-process pattern as the GST
 * `missed-books-layout-cache` — pruned lazily by TTL on each access.
 */
export interface BankLayoutCacheEntry {
  result: BankMatchResult;
  brs: BrsResult;
  periodLabel: string;
  periodStart: string;
  periodEnd: string;
  relativeTo?: string;
  layout: BankReconLayoutDto;
  sheetName: string;
  booksSheet: string;
  bankSheet: string;
  createdAt: number;
}

const TTL_MS = 30 * 60 * 1000;
const cache = new Map<string, BankLayoutCacheEntry>();
let counter = 0;

function pruneExpired(): void {
  const now = Date.now();
  for (const [id, entry] of cache) {
    if (now - entry.createdAt > TTL_MS) cache.delete(id);
  }
}

export function storeBankLayoutCache(entry: Omit<BankLayoutCacheEntry, 'createdAt'>): string {
  pruneExpired();
  const id = `bank_layout_cache_${++counter}_${Date.now().toString(36)}`;
  cache.set(id, { ...entry, createdAt: Date.now() });
  return id;
}

export function getBankLayoutCache(id: string): BankLayoutCacheEntry | undefined {
  pruneExpired();
  return cache.get(id);
}
