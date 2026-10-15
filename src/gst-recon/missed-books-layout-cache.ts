import { GstResultRow } from '../domain-tools/gst/match-passes';
import { GstReconLayoutDto } from './gst-recon.dto';

/**
 * Holds just enough of a completed CASUAL (missed_books_only) reconciliation run to
 * re-map its already-matched rows into a different sheet layout (categorized /
 * books_flat / portal_flat) without re-reading the workbook or re-running the matcher.
 * Scoped to casual runs only — the full/gated report (`mapToSheetActions`) has no
 * layout variants, so there is nothing to re-map there.
 */
export interface MissedBooksLayoutCacheEntry {
  portalLabel: string;
  booksLabel: string;
  resultRows: GstResultRow[];
  isSales: boolean;
  relativeTo?: string;
  /** The sheet name THIS run produced, for layout callers that want to detect "no-op: same layout requested again". */
  layout: GstReconLayoutDto;
  sheetName: string;
  createdAt: number;
}

// A short-lived, in-memory, single-process cache — same lifetime assumption as
// `pendingSheetCollision` on the client (chatReconChat.ts): it only needs to survive
// long enough for a same-session follow-up prompt, never across a server restart or
// between users. Entries are pruned lazily by TTL on each get/set rather than run on a
// timer, since recon runs are infrequent enough that a background sweep is unwarranted.
const TTL_MS = 30 * 60 * 1000; // 30 minutes — generous for "I'll decide on the layout in a bit"
const cache = new Map<string, MissedBooksLayoutCacheEntry>();
let counter = 0;

function pruneExpired(): void {
  const now = Date.now();
  for (const [id, entry] of cache) {
    if (now - entry.createdAt > TTL_MS) cache.delete(id);
  }
}

export function storeMissedBooksLayoutCache(
  entry: Omit<MissedBooksLayoutCacheEntry, 'createdAt'>,
): string {
  pruneExpired();
  const id = `layout_cache_${++counter}_${Date.now().toString(36)}`;
  cache.set(id, { ...entry, createdAt: Date.now() });
  return id;
}

export function getMissedBooksLayoutCache(id: string): MissedBooksLayoutCacheEntry | undefined {
  pruneExpired();
  return cache.get(id);
}
