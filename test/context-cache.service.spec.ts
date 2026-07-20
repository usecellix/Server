import { ContextCacheService } from '../src/excel-ai/services/context-cache.service';

describe('ContextCacheService (Spec 09 item 2)', () => {
  let cache: ContextCacheService;

  beforeEach(() => {
    cache = new ContextCacheService();
  });

  it('uses a 60-minute TTL', () => {
    expect(cache.getStats().ttlMs).toBe(60 * 60 * 1000);
  });

  it('hits when hash matches within TTL and counts hits', () => {
    cache.set('conv-1', 'toon-payload', 'prompt-A');
    expect(cache.get('conv-1', 'toon-payload')).toBe('prompt-A');
    expect(cache.get('conv-1', 'toon-payload')).toBe('prompt-A');
    expect(cache.getStats().hits).toBe(2);
    expect(cache.getStats().misses).toBe(0);
  });

  it('misses on hash mismatch', () => {
    cache.set('conv-1', 'toon-payload', 'prompt-A');
    expect(cache.get('conv-1', 'other-toon')).toBeNull();
    expect(cache.getStats().misses).toBe(1);
  });

  it('evicts oldest entries when over max size', () => {
    const max = cache.getStats().maxEntries;
    for (let i = 0; i < max + 5; i++) {
      cache.set(`conv-${i}`, `toon-${i}`, `prompt-${i}`);
    }
    expect(cache.getStats().size).toBe(max);
    expect(cache.get('conv-0', 'toon-0')).toBeNull();
    expect(cache.get(`conv-${max + 4}`, `toon-${max + 4}`)).toBe(`prompt-${max + 4}`);
  });
});
