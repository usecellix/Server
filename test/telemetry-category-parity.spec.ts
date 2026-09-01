import 'reflect-metadata';
import fs from 'fs';
import path from 'path';
import { TELEMETRY_CATEGORIES } from '../src/common/logging/dto/frontend-log-batch.dto';

/**
 * TASKS.md #159 — the frontend telemetry category list exists twice.
 *
 * `client/src/services/frontendTelemetry.ts` declares `FrontendTelemetryCategory`;
 * `frontend-log-batch.dto.ts` declares `CATEGORIES` and validates against it.
 * Adding `'verify'` to the client for TASKS.md #150 and not to the server made
 * every batch containing a verify event fail validation with 400 — and because
 * the DTO rejects the WHOLE batch, `accept.success` and the verification result
 * were dropped along with it. The apply had actually succeeded; the run merely
 * looked hung, with no error anywhere to explain it.
 *
 * This is `ARCHITECTURE.md` AD-7's drift pattern in a third place (after action
 * types and guard logic), so it gets the same treatment the action catalog got:
 * a test that reads the other side's source and fails when they disagree.
 */
describe('frontend telemetry category parity (TASKS.md #159)', () => {
  // The add-in lives in `frontend/`, not `client/` — this path was stale and the
  // suite had been failing on ENOENT, which meant this drift detector was itself
  // dead. Exactly the failure mode the docblock above warns about.
  const CLIENT_SOURCE = path.resolve(
    __dirname,
    '../../frontend/src/services/frontendTelemetry.ts',
  );

  function clientCategories(): string[] {
    const src = fs.readFileSync(CLIENT_SOURCE, 'utf8');
    const decl = /export type FrontendTelemetryCategory\s*=([\s\S]*?);/.exec(src);
    if (!decl) throw new Error('FrontendTelemetryCategory declaration not found');
    return [...decl[1].matchAll(/'([a-z]+)'/g)].map((m) => m[1]);
  }

  it('resolves the client source (a dead path would make this test vacuous)', () => {
    // TASKS.md #62's lesson: a cross-repo test whose import cannot resolve
    // fails silently in exactly the way the bug it guards does.
    expect(fs.existsSync(CLIENT_SOURCE)).toBe(true);
  });

  it('server accepts every category the client can emit', () => {
    const missing = clientCategories().filter((c) => !TELEMETRY_CATEGORIES.includes(c as never));
    expect(missing).toEqual([]);
  });

  it('includes the verify category the outcome verifier emits', () => {
    expect(TELEMETRY_CATEGORIES).toContain('verify');
    expect(clientCategories()).toContain('verify');
  });
});
