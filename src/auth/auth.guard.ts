import {
  CanActivate,
  ExecutionContext,
  Injectable,
  UnauthorizedException,
  createParamDecorator,
} from '@nestjs/common';
import { getAuth } from './auth';

export type AuthUserSession = NonNullable<
  Awaited<ReturnType<Awaited<ReturnType<typeof getAuth>>['api']['getSession']>>
>;

@Injectable()
export class AuthGuard implements CanActivate {
  async canActivate(context: ExecutionContext): Promise<boolean> {
    const request = context.switchToHttp().getRequest<{
      headers: Record<string, string | string[] | undefined>;
      userSession?: AuthUserSession;
    }>();

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
