import { Injectable, Logger, type MessageEvent } from '@nestjs/common';
import { Subject } from 'rxjs';

const TOKEN_TTL_MS = 5 * 60 * 1000;
/** How long Excel may call /excel-login/claim after the browser tab finishes sign-in. */
const CLAIM_TTL_MS = 2 * 60 * 1000;

export interface ExcelLoginClaim {
  sessionToken: string;
  userId: string;
  email: string | null;
}

/**
 * Pairs the Excel add-in's "waiting for email/password sign-in" SSE
 * connection (client/src/auth/useAuth.ts openEmailLoginPage) with the
 * browser tab that completes it (Landing-page/src/pages/LoginPage.tsx).
 *
 * Important: the Landing tab and the Excel task pane do **not** share a
 * cookie jar (system browser vs Office WebView2). So after notify() we keep a
 * short-lived claim the task pane redeems via POST /excel-login/claim to get
 * a Set-Cookie in the WebView — refetch alone is not enough.
 *
 * In-memory only — fine for a single Nest instance; a horizontally scaled
 * deployment would need this moved to a shared pub/sub (e.g. Redis).
 */
@Injectable()
export class ExcelLoginBridgeService {
  private readonly logger = new Logger(ExcelLoginBridgeService.name);
  private readonly pending = new Map<
    string,
    { subject: Subject<MessageEvent>; timeout: NodeJS.Timeout }
  >();
  private readonly claims = new Map<
    string,
    { claim: ExcelLoginClaim; timeout: NodeJS.Timeout }
  >();

  waitFor(token: string): Subject<MessageEvent> {
    this.clearWait(token);

    const subject = new Subject<MessageEvent>();
    const timeout = setTimeout(() => {
      subject.complete();
      this.pending.delete(token);
    }, TOKEN_TTL_MS);

    this.pending.set(token, { subject, timeout });
    return subject;
  }

  /**
   * Push login-complete to the waiting Excel SSE connection and retain a
   * one-time claim so the task pane can mint a WebView session cookie.
   */
  notify(token: string, claim: ExcelLoginClaim): boolean {
    this.storeClaim(token, claim);

    const entry = this.pending.get(token);
    if (!entry) {
      this.logger.debug(
        `No waiting Excel SSE for login token (expired or unknown) — claim retained for ${CLAIM_TTL_MS}ms`,
      );
      return false;
    }

    entry.subject.next({
      type: 'login-complete',
      data: JSON.stringify({ ok: true, email: claim.email }),
    });
    // Brief delay so Nest can flush the SSE event before the stream closes —
    // otherwise WebView2 EventSource often only sees onerror.
    setTimeout(() => {
      entry.subject.complete();
      this.clearWait(token);
    }, 75);
    return true;
  }

  /** One-time consume — returns null if missing/expired/already claimed. */
  consumeClaim(token: string): ExcelLoginClaim | null {
    const entry = this.claims.get(token);
    if (!entry) return null;
    clearTimeout(entry.timeout);
    this.claims.delete(token);
    return entry.claim;
  }

  private storeClaim(token: string, claim: ExcelLoginClaim): void {
    const existing = this.claims.get(token);
    if (existing) clearTimeout(existing.timeout);

    const timeout = setTimeout(() => {
      this.claims.delete(token);
    }, CLAIM_TTL_MS);
    this.claims.set(token, { claim, timeout });
  }

  private clearWait(token: string): void {
    const entry = this.pending.get(token);
    if (!entry) return;
    clearTimeout(entry.timeout);
    this.pending.delete(token);
  }
}
