import { Controller, Get, UseGuards } from '@nestjs/common';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { AuthGuard, Session, type AuthUserSession } from './auth.guard';

@Controller('auth')
export class AuthController {
  @Get('me')
  @UseGuards(AuthGuard)
  getProfile(@Session() session: AuthUserSession) {
    return {
      user: session.user,
      session: {
        id: session.session.id,
        expiresAt: session.session.expiresAt,
      },
    };
  }
}

/**
 * `providers` reflects what's actually USABLE right now, not what the code
 * supports — `email` and `google` are unconditional (email/password is
 * always on; Google's client id/secret have been set since launch), but
 * `microsoft` only appears when MICROSOFT_CLIENT_ID/SECRET are both set.
 * Without this, a provider with no credentials configured threw Better
 * Auth's CLIENT_ID_AND_SECRET_REQUIRED the moment someone clicked it — this
 * endpoint didn't exist at all, so CELLIX-landing-page's LoginPage always
 * rendered a "Continue with Microsoft" button that errored on click.
 *
 * Deliberately its OWN controller/prefix, NOT a route on AuthController
 * (`/auth/*`): CELLIX-landing-page/vercel.json's `/api/auth/:path*` rewrite
 * forwards WITH the `/api` prefix kept (that rule exists only for Better
 * Auth's own `/api/auth/*` catch-all). This app's `apiBaseUrl` is `/api` in
 * production, so `fetch('/auth/...')` becomes `/api/auth/...` on the wire —
 * which that rewrite sends to `api.usecellix.com/api/auth/...`, a path
 * `/auth/*`-prefixed routes never register (confirmed live: 404, even after
 * confirming the backend had restarted onto the right commit). This
 * controller's own `/sign-in-providers` prefix can't collide with that rule
 * no matter what apiBaseUrl is.
 */
@Controller('sign-in-providers')
export class SignInProvidersController {
  @Get()
  @SkipEnvelope()
  get() {
    const providers = ['email', 'google'];
    if (process.env.MICROSOFT_CLIENT_ID && process.env.MICROSOFT_CLIENT_SECRET) {
      providers.push('microsoft');
    }
    return { status: 'ok', providers };
  }
}
