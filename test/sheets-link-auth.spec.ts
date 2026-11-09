import { UnauthorizedException } from '@nestjs/common';
import { AppConfigService } from '../src/config/app-config.service';

const getSessionMock = jest.fn();
jest.mock('../src/auth/auth', () => ({
  getAuth: jest.fn().mockResolvedValue({ api: { getSession: (...args: unknown[]) => getSessionMock(...args) } }),
}));

import { AuthGuard } from '../src/auth/auth.guard';
import {
  hashSheetsKey,
  looksLikeSheetsKey,
  newSheetsKey,
  SHEETS_KEY_HEADER,
  SheetsKeyResolver,
} from '../src/sheets-link/sheets-link.key';

describe('Google Sheets link key authentication', () => {
  const config = { nodeEnv: 'production', evalBypassToken: '' } as unknown as AppConfigService;
  const ctx = (headers: Record<string, string>, request: Record<string, unknown> = {}) =>
    ({ switchToHttp: () => ({ getRequest: () => Object.assign(request, { headers }) }) }) as never;

  const KEY = newSheetsKey();
  beforeEach(() => {
    getSessionMock.mockReset();
    SheetsKeyResolver.resolve = async (hash) =>
      hash === hashSheetsKey(KEY) ? { userId: 'user-1', email: 'ca@firm.in', name: 'CA' } : null;
  });

  it('keys look right and only the hash is used for lookup', () => {
    expect(looksLikeSheetsKey(KEY)).toBe(true);
    expect(looksLikeSheetsKey('abc')).toBe(false);
    expect(hashSheetsKey(KEY)).toMatch(/^[0-9a-f]{64}$/);
    expect(hashSheetsKey(KEY)).not.toContain(KEY);
    expect(newSheetsKey()).not.toBe(KEY);
  });

  it('a valid key authenticates as its user without a browser session', async () => {
    const request: Record<string, unknown> = {};
    await expect(new AuthGuard(config).canActivate(ctx({ [SHEETS_KEY_HEADER]: KEY }, request))).resolves.toBe(true);
    expect((request.userSession as { user: { id: string } }).user.id).toBe('user-1');
    expect(getSessionMock).not.toHaveBeenCalled();
  });

  it('a wrong, malformed or revoked key is refused — never falls back to anonymous', async () => {
    for (const bad of [newSheetsKey(), 'cxs_short', 'not-a-key']) {
      await expect(new AuthGuard(config).canActivate(ctx({ [SHEETS_KEY_HEADER]: bad }))).rejects.toBeInstanceOf(UnauthorizedException);
    }
  });

  it('without the header the normal session check still runs', async () => {
    getSessionMock.mockResolvedValue(null);
    await expect(new AuthGuard(config).canActivate(ctx({}))).rejects.toBeInstanceOf(UnauthorizedException);
    expect(getSessionMock).toHaveBeenCalled();
  });
});
