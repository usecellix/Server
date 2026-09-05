import { UnauthorizedException } from '@nestjs/common';
import { AppConfigService } from '../src/config/app-config.service';

const getSessionMock = jest.fn();

jest.mock('../src/auth/auth', () => ({
  getAuth: jest.fn().mockResolvedValue({
    api: { getSession: (...args: unknown[]) => getSessionMock(...args) },
  }),
}));

// Imported after the mock so AuthGuard picks up the mocked getAuth.
import { AuthGuard, EVAL_BYPASS_HEADER } from '../src/auth/auth.guard';

/**
 * eval/run-live-eval.ts has no browser session — this is the narrow exception
 * that lets it authenticate. Must stay locked to non-production AND an
 * explicitly configured token; a bare unset token or a production nodeEnv
 * must fall through to the real session check exactly as before this existed.
 */
describe('AuthGuard — eval bypass', () => {
  function buildContext(headers: Record<string, string>) {
    return {
      switchToHttp: () => ({
        getRequest: () => ({ headers }),
      }),
    } as never;
  }

  beforeEach(() => {
    getSessionMock.mockReset();
  });

  it('bypasses the real session check when nodeEnv is not production and the header matches the configured token', async () => {
    const config = {
      nodeEnv: 'development',
      evalBypassToken: 'secret-eval-token',
    } as unknown as AppConfigService;
    const guard = new AuthGuard(config);

    const allowed = await guard.canActivate(
      buildContext({ [EVAL_BYPASS_HEADER]: 'secret-eval-token' }),
    );

    expect(allowed).toBe(true);
    expect(getSessionMock).not.toHaveBeenCalled();
  });

  it('falls through to the real session check when the header does not match', async () => {
    const config = {
      nodeEnv: 'development',
      evalBypassToken: 'secret-eval-token',
    } as unknown as AppConfigService;
    const guard = new AuthGuard(config);
    getSessionMock.mockResolvedValue(null);

    await expect(
      guard.canActivate(buildContext({ [EVAL_BYPASS_HEADER]: 'wrong-token' })),
    ).rejects.toBeInstanceOf(UnauthorizedException);
    expect(getSessionMock).toHaveBeenCalled();
  });

  it('falls through to the real session check when no bypass token is configured at all', async () => {
    const config = {
      nodeEnv: 'development',
      evalBypassToken: undefined,
    } as unknown as AppConfigService;
    const guard = new AuthGuard(config);
    getSessionMock.mockResolvedValue(null);

    await expect(
      guard.canActivate(buildContext({ [EVAL_BYPASS_HEADER]: 'anything' })),
    ).rejects.toBeInstanceOf(UnauthorizedException);
    expect(getSessionMock).toHaveBeenCalled();
  });

  it('never bypasses in production, even with a matching token and header', async () => {
    const config = {
      nodeEnv: 'production',
      evalBypassToken: 'secret-eval-token',
    } as unknown as AppConfigService;
    const guard = new AuthGuard(config);
    getSessionMock.mockResolvedValue(null);

    await expect(
      guard.canActivate(buildContext({ [EVAL_BYPASS_HEADER]: 'secret-eval-token' })),
    ).rejects.toBeInstanceOf(UnauthorizedException);
    expect(getSessionMock).toHaveBeenCalled();
  });

  it('still authenticates normally via the real session when no bypass header is sent', async () => {
    const config = {
      nodeEnv: 'development',
      evalBypassToken: 'secret-eval-token',
    } as unknown as AppConfigService;
    const guard = new AuthGuard(config);
    getSessionMock.mockResolvedValue({
      user: { id: 'u1' },
      session: { id: 's1' },
    });

    const allowed = await guard.canActivate(buildContext({}));

    expect(allowed).toBe(true);
    expect(getSessionMock).toHaveBeenCalled();
  });
});
