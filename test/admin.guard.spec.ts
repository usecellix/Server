import { ExecutionContext, UnauthorizedException } from '@nestjs/common';
import { AdminGuard, ADMIN_TOKEN_HEADER } from '../src/admin/admin.guard';

function contextWith(headers: Record<string, string | string[] | undefined>): ExecutionContext {
  return {
    switchToHttp: () => ({ getRequest: () => ({ headers }) }),
  } as unknown as ExecutionContext;
}

describe('AdminGuard', () => {
  it('refuses everything when no admin token is configured, not just missing headers', () => {
    const guard = new AdminGuard({ adminApiToken: undefined } as never);
    expect(() => guard.canActivate(contextWith({ [ADMIN_TOKEN_HEADER]: 'anything' }))).toThrow(UnauthorizedException);
  });

  it('rejects a request with no token header', () => {
    const guard = new AdminGuard({ adminApiToken: 'secret' } as never);
    expect(() => guard.canActivate(contextWith({}))).toThrow(UnauthorizedException);
  });

  it('rejects a request with the wrong token', () => {
    const guard = new AdminGuard({ adminApiToken: 'secret' } as never);
    expect(() => guard.canActivate(contextWith({ [ADMIN_TOKEN_HEADER]: 'wrong' }))).toThrow(UnauthorizedException);
  });

  it('allows a request with the matching token', () => {
    const guard = new AdminGuard({ adminApiToken: 'secret' } as never);
    expect(guard.canActivate(contextWith({ [ADMIN_TOKEN_HEADER]: 'secret' }))).toBe(true);
  });

  it('handles a duplicated header (array) by checking the first value', () => {
    const guard = new AdminGuard({ adminApiToken: 'secret' } as never);
    expect(guard.canActivate(contextWith({ [ADMIN_TOKEN_HEADER]: ['secret', 'other'] }))).toBe(true);
  });
});
