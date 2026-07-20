import { MongoClient } from 'mongodb';
import { importEsm } from './import-esm';

export type BetterAuthInstance = {
  handler: (request: Request) => Promise<Response>;
  api: {
    getSession: (args: { headers: Headers }) => Promise<{
      user: {
        id: string;
        name: string;
        email: string;
        image?: string | null;
        emailVerified: boolean;
        createdAt: Date;
        updatedAt: Date;
      };
      session: {
        id: string;
        userId: string;
        expiresAt: Date;
        token: string;
        createdAt: Date;
        updatedAt: Date;
      };
    } | null>;
  };
};

let authPromise: Promise<BetterAuthInstance> | null = null;

/** Standalone MongoDB (no replica set) rejects default retryable writes used by the driver. */
function withRetryWritesDisabled(uri: string): string {
  if (/[?&]retryWrites=/i.test(uri)) {
    return uri.replace(/([?&]retryWrites=)[^&]*/i, '$1false');
  }
  return uri.includes('?') ? `${uri}&retryWrites=false` : `${uri}?retryWrites=false`;
}

export function getAuth(): Promise<BetterAuthInstance> {
  if (!authPromise) {
    authPromise = createAuth();
  }
  return authPromise;
}

async function createAuth(): Promise<BetterAuthInstance> {
  const { betterAuth } = await importEsm<{
    betterAuth: (config: Record<string, unknown>) => BetterAuthInstance;
  }>('better-auth');
  const { mongodbAdapter } = await importEsm<{
    mongodbAdapter: (
      db: ReturnType<MongoClient['db']>,
      options?: { client: MongoClient },
    ) => unknown;
  }>('better-auth/adapters/mongodb');

  const mongoUrl = withRetryWritesDisabled(
    process.env.MONGODB_URL || 'mongodb://127.0.0.1:27017/cellix',
  );
  const dbName = process.env.MONGODB_DB_NAME || 'cellix';
  const clientOrigin = process.env.CLIENT_ORIGIN || 'https://localhost:3000';
  const betterAuthUrl = process.env.BETTER_AUTH_URL || clientOrigin;

  // Task pane is https://localhost:3000; Google console may still list http — allow both.
  const trustedOrigins = Array.from(
    new Set([
      clientOrigin.replace(/\/$/, ''),
      betterAuthUrl.replace(/\/$/, ''),
      'https://localhost:3000',
      'http://localhost:3000',
    ]),
  );

  const client = new MongoClient(mongoUrl);
  await client.connect();
  const db = client.db(dbName);

  return betterAuth({
    database: mongodbAdapter(db),
    baseURL: betterAuthUrl,
    secret: process.env.BETTER_AUTH_SECRET,
    trustedOrigins,
    socialProviders: {
      google: {
        clientId: process.env.GOOGLE_CLIENT_ID as string,
        clientSecret: process.env.GOOGLE_CLIENT_SECRET as string,
        // Always show Google's account chooser when cookies exist in this WebView.
        prompt: 'select_account',
      },
      microsoft: {
        clientId: process.env.MICROSOFT_CLIENT_ID as string,
        clientSecret: process.env.MICROSOFT_CLIENT_SECRET as string,
        tenantId: process.env.MICROSOFT_TENANT_ID || 'common',
        prompt: 'select_account',
      },
    },
    account: {
      accountLinking: {
        enabled: true,
        trustedProviders: ['google', 'microsoft'],
      },
      // Excel WebView starts OAuth; Google often finishes in another browser jar.
      // Without this, Better Auth throws state_security_mismatch (state cookie missing).
      // DB state verification still runs — see:
      // https://better-auth.com/docs/reference/errors/state_mismatch
      skipStateCookieCheck: true,
      storeStateStrategy: 'database',
    },
    advanced: {
      defaultCookieAttributes: {
        // Same-site cookies work for same-origin Vite proxy callbacks.
        sameSite: 'lax',
        secure: betterAuthUrl.startsWith('https'),
        httpOnly: true,
        path: '/',
      },
    },
  });
}
