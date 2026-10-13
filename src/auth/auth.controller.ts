import { Controller, Get, UseGuards } from '@nestjs/common';
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

  /**
   * `providers` reflects what's actually USABLE right now, not what the
   * code supports — `email` and `google` are unconditional (email/password
   * is always on; Google's client id/secret have been set since launch),
   * but `microsoft` only appears when MICROSOFT_CLIENT_ID/SECRET are both
   * set. Without this, a provider with no credentials configured threw
   * CLIENT_ID_AND_SECRET_REQUIRED (Better Auth) the moment someone clicked
   * it — this endpoint existed and unconditionally claimed `microsoft` was
   * available, which is what let that ship. CELLIX-landing-page's
   * LoginPage reads this to hide a provider button rather than show one
   * that errors on click.
   */
  @Get('health')
  authHealth() {
    const providers = ['email', 'google'];
    if (process.env.MICROSOFT_CLIENT_ID && process.env.MICROSOFT_CLIENT_SECRET) {
      providers.push('microsoft');
    }
    return { status: 'ok', providers };
  }
}
