// cellix_backend/src/excel-ai/services/context-cache.service.ts

import { Injectable, Logger } from '@nestjs/common';
import { createHash } from 'crypto';

/** Spec 09 item 2: TTL bounds abandoned conversations only — hash still owns correctness. */
const CONTEXT_CACHE_TTL_MS = 60 * 60 * 1000; // was 10 min
/** Soft cap so widening TTL cannot grow memory without bound. */
const CONTEXT_CACHE_MAX_ENTRIES = 100;

/**
 * Caches the expensive SheetAnalyzer + TOON output per conversation turn.
 *
 * Strategy: hash the TOON-compressed string. If hash matches last turn's hash
 * for this conversationId, return the cached promptContext instead of re-running.
 */
@Injectable()
export class ContextCacheService {
  private readonly logger = new Logger(ContextCacheService.name);

  private readonly memCache = new Map<
    string,
    { hash: string; promptContext: string; timestamp: number }
  >();

  private readonly TTL_MS = CONTEXT_CACHE_TTL_MS;
  private readonly MAX_ENTRIES = CONTEXT_CACHE_MAX_ENTRIES;

  private hits = 0;
  private misses = 0;

  /**
   * Check if we have a valid cached promptContext for this conversation + sheet state.
   */
  get(conversationId: string, toonPayload: string): string | null {
    const hash = this.hashToon(toonPayload);
    const cached = this.memCache.get(conversationId);

    if (!cached) {
      this.misses += 1;
      return null;
    }
    if (cached.hash !== hash) {
      this.misses += 1;
      return null;
    }
    if (Date.now() - cached.timestamp > this.TTL_MS) {
      this.memCache.delete(conversationId);
      this.misses += 1;
      return null;
    }

    // Refresh LRU order (Map insertion order).
    this.memCache.delete(conversationId);
    this.memCache.set(conversationId, { ...cached, timestamp: cached.timestamp });

    this.hits += 1;
    this.logger.debug(
      `Context cache hit for conversation ${conversationId} (hits=${this.hits} misses=${this.misses})`,
    );
    return cached.promptContext;
  }

  /**
   * Store a newly-built promptContext in the cache.
   */
  set(conversationId: string, toonPayload: string, promptContext: string): void {
    const hash = this.hashToon(toonPayload);
    if (this.memCache.has(conversationId)) {
      this.memCache.delete(conversationId);
    }
    this.memCache.set(conversationId, {
      hash,
      promptContext,
      timestamp: Date.now(),
    });
    this.evictIfNeeded();
    this.logger.debug(`Context cache stored for conversation ${conversationId}`);
  }

  /**
   * Invalidate cache when an action is applied (sheet has changed).
   */
  invalidate(conversationId: string): void {
    this.memCache.delete(conversationId);
    this.logger.debug(`Context cache invalidated for conversation ${conversationId}`);
  }

  /** Exposed for ops / tests — hit rate over process lifetime. */
  getStats(): { hits: number; misses: number; size: number; ttlMs: number; maxEntries: number } {
    return {
      hits: this.hits,
      misses: this.misses,
      size: this.memCache.size,
      ttlMs: this.TTL_MS,
      maxEntries: this.MAX_ENTRIES,
    };
  }

  private evictIfNeeded(): void {
    while (this.memCache.size > this.MAX_ENTRIES) {
      const oldestKey = this.memCache.keys().next().value as string | undefined;
      if (!oldestKey) break;
      this.memCache.delete(oldestKey);
      this.logger.debug(`Context cache LRU evicted conversation ${oldestKey}`);
    }
  }

  private hashToon(toonPayload: string): string {
    return createHash('sha256').update(toonPayload).digest('hex').slice(0, 16);
  }
}
