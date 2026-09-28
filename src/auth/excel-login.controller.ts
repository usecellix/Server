import {
  BadRequestException,
  Body,
  Controller,
  Post,
  Query,
  Res,
  Sse,
  UnauthorizedException,
  UseGuards,
  type MessageEvent,
} from '@nestjs/common';
import type { FastifyReply } from 'fastify';
import type { Observable } from 'rxjs';
import { AuthGuard, Session, type AuthUserSession } from './auth.guard';
import { ExcelLoginBridgeService } from './excel-login-bridge.service';
import { importEsm } from './import-esm';

/** Matches the token client/src/auth/useAuth.ts generates with crypto.randomUUID(). */
const TOKEN_PATTERN = /^[a-zA-Z0-9-]{8,100}$/;

/**
 * Deliberately NOT under /auth — that prefix is reserved for Better Auth's
 * own /api/auth/* catch-all (auth-route.registrar.ts) plus this app's own
 * /auth/* routes (AuthController), and client/vite.config.ts's dev proxy
 * treats /api/auth/* as a single kept-prefix rule pointed at Better Auth.
 * A route here under /auth would collide with that rule and 404 through the
 * proxy even though it exists on the Nest side — see excel-login/*'s own
 * proxy rule (the general prefix-stripped /api/* one) instead.
 */
@Controller('excel-login')
export class ExcelLoginController {
  constructor(private readonly excelLoginBridge: ExcelLoginBridgeService) {}

  /**
   * The Excel add-in opens this once it launches the email/password login
   * page in an external browser tab (client/src/auth/useAuth.ts
   * openEmailLoginPage) and waits here instead of polling get-session. Push,
   * not poll — see complete() / claim() below for the other half.
   */
  @Sse('wait')
  wait(@Query('token') token: string): Observable<MessageEvent> {
    if (!token || !TOKEN_PATTERN.test(token)) {
      throw new BadRequestException('Invalid or missing token');
    }
    return this.excelLoginBridge.waitFor(token);
  }

  /**
   * The browser tab (Landing-page/src/pages/LoginPage.tsx) calls this right
   * after a successful sign-in/sign-up, authenticated by its own session
   * cookie — proves the token's owner actually completed a real login before
   * any Excel task pane is told to switch to the chat screen.
   */
  @Post('complete')
  @UseGuards(AuthGuard)
  complete(@Body('token') token: string, @Session() session: AuthUserSession) {
    if (!token || !TOKEN_PATTERN.test(token)) {
      throw new BadRequestException('Invalid or missing token');
    }
    const delivered = this.excelLoginBridge.notify(token, {
      sessionToken: session.session.token,
      userId: session.user.id,
      email: session.user.email ?? null,
    });
    return { delivered };
  }

  /**
   * Called by the Excel task pane after it receives the SSE login-complete
   * push. Sets Better Auth's session cookie on *this* response so the Office
   * WebView2 jar (which never saw the Landing-tab login) can pass get-session
   * and AuthGate flips to chat.
   */
  @Post('claim')
  async claim(
    @Body('token') token: string,
    @Res({ passthrough: true }) reply: FastifyReply,
  ) {
    if (!token || !TOKEN_PATTERN.test(token)) {
      throw new BadRequestException('Invalid or missing token');
    }
    const claim = this.excelLoginBridge.consumeClaim(token);
    if (!claim?.sessionToken) {
      throw new UnauthorizedException('Login claim expired or already used');
    }

    const setCookie = await buildSessionSetCookie(claim.sessionToken);
    // Array form — Fastify would overwrite if header() was called twice.
    reply.header('set-cookie', [
      setCookie,
      // Clear legacy __Secure- name from earlier cookie configs.
      '__Secure-better-auth.session_token=; Max-Age=0; Path=/; HttpOnly; Secure; SameSite=Lax',
    ]);
    return { ok: true, email: claim.email };
  }
}

/**
 * Produce the same signed `better-auth.session_token` Set-Cookie Better Auth
 * would issue on sign-in, so get-session accepts it.
 */
async function buildSessionSetCookie(sessionToken: string): Promise<string> {
  const secret = process.env.BETTER_AUTH_SECRET;
  if (!secret) {
    throw new BadRequestException('BETTER_AUTH_SECRET is not configured');
  }

  const clientOrigin = process.env.CLIENT_ORIGIN || 'https://localhost:3000';
  const betterAuthUrl = process.env.BETTER_AUTH_URL || clientOrigin;

  const { getCookies } = await importEsm<{
    getCookies: (options: Record<string, unknown>) => {
      sessionToken: { name: string; attributes: Record<string, unknown> };
    };
  }>('better-auth/cookies');

  const { serializeSignedCookie } = await importEsm<{
    serializeSignedCookie: (
      key: string,
      value: string,
      secret: string,
      opt?: Record<string, unknown>,
    ) => Promise<string>;
  }>('better-call');

  const cookies = getCookies({
    baseURL: betterAuthUrl,
    advanced: {
      useSecureCookies: process.env.NODE_ENV === 'production',
      defaultCookieAttributes: {
        sameSite: 'lax',
        secure: process.env.NODE_ENV === 'production',
        httpOnly: true,
        path: '/',
      },
    },
  });

  return serializeSignedCookie(
    cookies.sessionToken.name,
    sessionToken,
    secret,
    {
      ...cookies.sessionToken.attributes,
      maxAge: 60 * 60 * 24 * 7,
    },
  );
}
