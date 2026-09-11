import {
  CanActivate,
  ExecutionContext,
  Injectable,
  UnauthorizedException,
  createParamDecorator,
} from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';
import { getAuth } from './auth';

export type AuthUserSession = NonNullable<
  Awaited<ReturnType<Awaited<ReturnType<typeof getAuth>>['api']['getSession']>>
>;

/** Header eval/run-live-eval.ts sends to authenticate without a browser session. */
export const EVAL_BYPASS_HEADER = 'x-cellix-eval-bypass';

@Injectable()
export class AuthGuard implements CanActivate {
  constructor(private readonly config: AppConfigService) {}

  async canActivate(context: ExecutionContext): Promise<boolean> {
    const request = context.switchToHttp().getRequest<{
      headers: Record<string, string | string[] | undefined>;
      userSession?: AuthUserSession;
    }>();

    // eval/run-live-eval.ts has no browser session to send — this exists so it
    // can exercise the real Router -> Tier 0-3 pipeline without one. Locked to
    // non-production AND an explicitly configured token, so a bare unset stays
    // exactly as strict as before (no header the harness sends can pass unless
    // the operator has also opted in server-side). TASKS.md model-swap-eval.
    const bypassToken = this.config.evalBypassToken;
    if (this.config.nodeEnv !== 'production' && bypassToken) {
      const header = request.headers[EVAL_BYPASS_HEADER];
      const presented = Array.isArray(header) ? header[0] : header;
      if (presented && presented === bypassToken) {
        return true;
      }
    }

    const auth = await getAuth();
    const headers = toWebHeaders(request.headers);
    const session = await auth.api.getSession({ headers });

    if (!session) {
      throw new UnauthorizedException('Authentication required');
    }

    request.userSession = session;
    return true;
  }
}

export const Session = createParamDecorator(
  (_data: unknown, context: ExecutionContext): AuthUserSession | undefined => {
    const request = context.switchToHttp().getRequest<{ userSession?: AuthUserSession }>();
    return request.userSession;
  },
);

function toWebHeaders(nodeHeaders: Record<string, string | string[] | undefined>): Headers {
  const headers = new Headers();
  for (const [key, value] of Object.entries(nodeHeaders)) {
    if (value === undefined) continue;
    if (Array.isArray(value)) {
      for (const entry of value) headers.append(key, entry);
    } else {
      headers.set(key, value);
    }
  }
  return headers;
}
