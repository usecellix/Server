// cellix_backend/src/excel-ai/services/context-cache.service.ts

import { Injectable, Logger } from '@nestjs/common';
import { createHash } from 'crypto';

/**
 * Caches the expensive SheetAnalyzer + TOON output per conversation turn.
 *
 * Problem solved: every follow-up message re-runs SheetAnalyzerService,
 * formula enrichment, and TOON compression even when the sheet hasn't changed.
 *
 * Strategy: hash the TOON-compressed string. If hash matches last turn's hash
 * for this conversationId, return the cached promptContext instead of re-running.
 */
@Injectable()
export class ContextCacheService {
  private readonly logger = new Logger(ContextCacheService.name);

  // In-memory cache for the current process (first line of defense)
  // Key: conversationId, Value: { hash, promptContext, timestamp }
  private readonly memCache = new Map<
    string,
    { hash: string; promptContext: string; timestamp: number }
  >();

  // TTL: 10 minutes. After this, always re-run analysis.
  private readonly TTL_MS = 10 * 60 * 1000;

  /**
   * Check if we have a valid cached promptContext for this conversation + sheet state.
   *
   * @param conversationId  The active conversation ID
   * @param toonPayload     The raw TOON-compressed string from the frontend
   * @returns               Cached promptContext if valid, null if must re-analyze
   */
  get(conversationId: string, toonPayload: string): string | null {
    const hash = this.hashToon(toonPayload);
    const cached = this.memCache.get(conversationId);

    if (!cached) return null;
    if (cached.hash !== hash) return null;
    if (Date.now() - cached.timestamp > this.TTL_MS) {
      this.memCache.delete(conversationId);
      return null;
    }

    this.logger.debug(`Context cache hit for conversation ${conversationId}`);
    return cached.promptContext;
  }

  /**
   * Store a newly-built promptContext in the cache.
   */
  set(conversationId: string, toonPayload: string, promptContext: string): void {
    const hash = this.hashToon(toonPayload);
    this.memCache.set(conversationId, {
      hash,
      promptContext,
      timestamp: Date.now(),
    });
    this.logger.debug(`Context cache stored for conversation ${conversationId}`);
  }

  /**
   * Invalidate cache when an action is applied (sheet has changed).
   * Call this after POST /audit/apply/:changeSetId.
   */
  invalidate(conversationId: string): void {
    this.memCache.delete(conversationId);
    this.logger.debug(`Context cache invalidated for conversation ${conversationId}`);
  }

  /**
   * Compute a short hash of the TOON string.
   * SHA-256 truncated to 16 chars — fast and collision-resistant enough.
   */
  private hashToon(toonPayload: string): string {
    return createHash('sha256').update(toonPayload).digest('hex').slice(0, 16);
  }
}
