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
  // This path has flip-flopped between `frontend/` and `client/` across past
  // sessions/checkouts (TASKS.md #166, #176) — the add-in repo has been
  // checked out under both names at different times. As of TASKS.md #221/#227
  // (Sept 2026), the actual checkout on disk is `client/` (git remote
  // usecellix/client), verified against this filesystem, not assumed. If this
  // starts failing existence again, check `ls` next to `Server/` before
  // guessing which name is current — that is exactly what silenced this
  // detector the first two times. TASKS.md #166, #176.
  const CLIENT_SOURCE = path.resolve(
    __dirname,
    '../../client/src/services/frontendTelemetry.ts',
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
