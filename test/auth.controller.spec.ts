import { AuthController } from '../src/auth/auth.controller';

describe('AuthController.authHealth', () => {
  const ORIGINAL_ENV = process.env;

  beforeEach(() => {
    process.env = { ...ORIGINAL_ENV };
  });

  afterAll(() => {
    process.env = ORIGINAL_ENV;
  });

  it('omits microsoft when MICROSOFT_CLIENT_ID/SECRET are not both set (the live-prod state that was previously misreported)', () => {
    delete process.env.MICROSOFT_CLIENT_ID;
    delete process.env.MICROSOFT_CLIENT_SECRET;
    const controller = new AuthController();
    expect(controller.authHealth()).toEqual({ status: 'ok', providers: ['email', 'google'] });
  });

  it('omits microsoft when only the client id is set', () => {
    process.env.MICROSOFT_CLIENT_ID = 'some-id';
    delete process.env.MICROSOFT_CLIENT_SECRET;
    const controller = new AuthController();
    expect(controller.authHealth().providers).not.toContain('microsoft');
  });

  it('includes microsoft once both client id and secret are set', () => {
    process.env.MICROSOFT_CLIENT_ID = 'some-id';
    process.env.MICROSOFT_CLIENT_SECRET = 'some-secret';
    const controller = new AuthController();
    expect(controller.authHealth()).toEqual({ status: 'ok', providers: ['email', 'google', 'microsoft'] });
  });
});
