import { CanActivate, ExecutionContext, Injectable, UnauthorizedException } from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';

export const ADMIN_TOKEN_HEADER = 'x-cellix-admin-token';

/**
 * Gates `/admin/*` — server-to-server only, never a browser session. The
 * Dashboard is its own app with its own password-gated login
 * (`Dashboard/src/lib/session.ts`); once an operator is in there, its Next.js
 * server calls this API with one shared secret, the same shape
 * `EVAL_BYPASS_HEADER` already uses in `AuthGuard` for the eval harness.
 * Deliberately NOT `AuthGuard`/Better Auth — an admin operator isn't a
 * product user, there is nothing to look up a `userSession` for, and reusing
 * user auth would mean "is this Google/Microsoft account an admin" becomes a
 * second place account privilege has to be checked correctly.
 */
@Injectable()
export class AdminGuard implements CanActivate {
  constructor(private readonly config: AppConfigService) {}

  canActivate(context: ExecutionContext): boolean {
    const token = this.config.adminApiToken;
    if (!token) {
      // Unconfigured means the admin API is off, not open — never fall
      // through to "any request passes" for a missing secret.
      throw new UnauthorizedException('ADMIN_API_NOT_CONFIGURED');
    }

    const request = context.switchToHttp().getRequest<{
      headers: Record<string, string | string[] | undefined>;
    }>();
    const header = request.headers[ADMIN_TOKEN_HEADER];
    const presented = Array.isArray(header) ? header[0] : header;
    if (!presented || presented !== token) {
      throw new UnauthorizedException('Admin authentication required');
    }
    return true;
  }
}
