// cellix_backend/src/common/cache/context-cache.service.ts
//
// Moved from excel-ai/services/ (TASKS.md #74) — PerformanceMetricsService
// (audit/) needs read access to getStats() for cache-hit-rate reporting, and
// AuditModule is already imported BY ExcelAiModule, so ExcelAiModule cannot
// import AuditModule back without a circular module dependency. This service
// has no DB dependency and no other coupling to excel-ai/ internals, making
// it a clean move to a neutral shared location both modules can import.

import { Injectable, Logger } from '@nestjs/common';
import { createHash } from 'crypto';

/** Spec 09 item 2: TTL bounds abandoned conversations only — hash still owns correctness. */
const CONTEXT_CACHE_TTL_MS = 60 * 60 * 1000; // was 10 min
/** Stable workbook state cache: 24h TTL for reuse across conversations. */
const STABLE_CACHE_TTL_MS = 24 * 60 * 60 * 1000;
/** Soft cap so widening TTL cannot grow memory without bound. */
const CONTEXT_CACHE_MAX_ENTRIES = 100;
/** Separate cap for stable cache — higher because it lives longer and serves many conversations. */
const STABLE_CACHE_MAX_ENTRIES = 50;

/**
 * Caches the expensive SheetAnalyzer + TOON output per conversation turn.
 *
 * Strategy: hash the TOON-compressed string. If hash matches last turn's hash
 * for this conversationId, return the cached promptContext instead of re-running.
 *
 * TASKS.md #68 — extended with a second, stable cache layer keyed by toonHash
 * alone (not conversationId), so a workbook in a stable state gets prompt-context
 * reuse ACROSS conversations, not just within one — a fresh conversation on an
 * already-seen sheet state no longer pays the full rebuild cost.
 */
@Injectable()
export class ContextCacheService {
  private readonly logger = new Logger(ContextCacheService.name);

  private readonly memCache = new Map<
    string,
    { hash: string; promptContext: string; timestamp: number }
  >();

  /** Stable cache keyed by toonHash (cross-conversation reuse). */
  private readonly stableCache = new Map<
    string,
    { promptContext: string; timestamp: number }
  >();

  private readonly TTL_MS = CONTEXT_CACHE_TTL_MS;
  private readonly MAX_ENTRIES = CONTEXT_CACHE_MAX_ENTRIES;
  private readonly STABLE_TTL_MS = STABLE_CACHE_TTL_MS;
  private readonly STABLE_MAX_ENTRIES = STABLE_CACHE_MAX_ENTRIES;

  private hits = 0;
  private misses = 0;
  private stableHits = 0;
  private stableMisses = 0;

  /**
   * Check if we have a valid cached promptContext for this conversation + sheet state.
   * Falls back to stable cache (cross-conversation) if conversation-level cache misses.
   */
  get(conversationId: string, toonPayload: string): string | null {
    const hash = this.hashToon(toonPayload);
    const cached = this.memCache.get(conversationId);

    // Try conversation-scoped cache first
    if (cached && cached.hash === hash && Date.now() - cached.timestamp <= this.TTL_MS) {
      // Refresh LRU order (Map insertion order).
      this.memCache.delete(conversationId);
      this.memCache.set(conversationId, { ...cached, timestamp: cached.timestamp });

      this.hits += 1;
      this.logger.debug(
        `Context cache hit (conversation) ${conversationId} (hits=${this.hits} misses=${this.misses})`,
      );
      return cached.promptContext;
    }

    // Clean up stale conversation entry if present
    if (cached) {
      this.memCache.delete(conversationId);
    }

    // Fall back to stable cache (cross-conversation, same sheet state)
    const stableCached = this.stableCache.get(hash);
    if (stableCached && Date.now() - stableCached.timestamp <= this.STABLE_TTL_MS) {
      // Refresh LRU order
      this.stableCache.delete(hash);
      this.stableCache.set(hash, { ...stableCached, timestamp: Date.now() });

      this.stableHits += 1;
      this.logger.debug(
        `Context cache hit (stable, workbook-wide) for hash ${hash} (hits=${this.stableHits} misses=${this.stableMisses})`,
      );

      // Populate conversation-scoped cache for faster future hits
      this.memCache.set(conversationId, {
        hash,
        promptContext: stableCached.promptContext,
        timestamp: Date.now(),
      });

      return stableCached.promptContext;
    }

    // Cache miss
    this.misses += 1;
    if (!stableCached) {
      this.stableMisses += 1;
    }
    return null;
  }

  /**
   * Store a newly-built promptContext in both conversation and stable caches.
   */
  set(conversationId: string, toonPayload: string, promptContext: string): void {
    const hash = this.hashToon(toonPayload);
    const now = Date.now();

    // Store in conversation-scoped cache
    if (this.memCache.has(conversationId)) {
      this.memCache.delete(conversationId);
    }
    this.memCache.set(conversationId, {
      hash,
      promptContext,
      timestamp: now,
    });

    // Store in stable cache (for cross-conversation reuse)
    if (this.stableCache.has(hash)) {
      this.stableCache.delete(hash);
    }
    this.stableCache.set(hash, {
      promptContext,
      timestamp: now,
    });

    this.evictIfNeeded();
    this.logger.debug(
      `Context cache stored for conversation ${conversationId} (hash: ${hash.slice(0, 8)}...)`,
    );
  }

  /**
   * Invalidate cache when an action is applied (sheet has changed).
   * Only invalidates conversation-scoped; stable cache is time-scoped only.
   */
  invalidate(conversationId: string): void {
    this.memCache.delete(conversationId);
    this.logger.debug(`Context cache invalidated for conversation ${conversationId}`);
  }

  /** Exposed for ops / tests — hit rate over process lifetime (both caches). */
  getStats(): {
    hits: number;
    misses: number;
    size: number;
    stableHits: number;
    stableMisses: number;
    stableSize: number;
    ttlMs: number;
    maxEntries: number;
    stableTtlMs: number;
    stableMaxEntries: number;
  } {
    return {
      hits: this.hits,
      misses: this.misses,
      size: this.memCache.size,
      stableHits: this.stableHits,
      stableMisses: this.stableMisses,
      stableSize: this.stableCache.size,
      ttlMs: this.TTL_MS,
      maxEntries: this.MAX_ENTRIES,
      stableTtlMs: this.STABLE_TTL_MS,
      stableMaxEntries: this.STABLE_MAX_ENTRIES,
    };
  }

  private evictIfNeeded(): void {
    // Evict from conversation cache if needed
    while (this.memCache.size > this.MAX_ENTRIES) {
      const oldestKey = this.memCache.keys().next().value as string | undefined;
      if (!oldestKey) break;
      this.memCache.delete(oldestKey);
      this.logger.debug(`Context cache LRU evicted conversation ${oldestKey}`);
    }

    // Evict from stable cache if needed
    while (this.stableCache.size > this.STABLE_MAX_ENTRIES) {
      const oldestHash = this.stableCache.keys().next().value as string | undefined;
      if (!oldestHash) break;
      this.stableCache.delete(oldestHash);
      this.logger.debug(`Stable cache LRU evicted (hash: ${oldestHash.slice(0, 8)}...)`);
    }
  }

  private hashToon(toonPayload: string): string {
    return createHash('sha256').update(toonPayload).digest('hex').slice(0, 16);
  }
}
