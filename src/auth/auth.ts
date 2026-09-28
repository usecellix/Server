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
  const marketingSiteOrigin = process.env.MARKETING_SITE_ORIGIN || '';
  // Marketing site (Landing-page) — hosts /login and /register, which the
  // Excel add-in opens in an external browser tab for email/password
  // (client/src/auth/useAuth.ts openEmailLoginPage). Without this, Better
  // Auth rejects sign-in/sign-up requests made from that origin.
  const marketingSiteOrigins = (process.env.MARKETING_SITE_ORIGIN || '')
    .split(',')
    .map((origin) => origin.trim().replace(/\/$/, ''))
    .filter(Boolean);

  // Task pane is https://localhost:3000; Google console may still list http — allow both.
  // MARKETING_SITE_ORIGIN is comma-separated (Vite picks the next free port, so both 5173
  // and 5174 may be active during dev). Split and normalize each one.
  const origins = [
    clientOrigin.replace(/\/$/, ''),
    betterAuthUrl.replace(/\/$/, ''),
    'https://localhost:3000',
    'http://localhost:3000',
  ];
  if (marketingSiteOrigin) {
    origins.push(
        ...marketingSiteOrigins,
    ...marketingSiteOrigin.split(',').map((o) => o.trim().replace(/\/$/, '')),
    );
  }
  const trustedOrigins = Array.from(new Set(origins));

  const client = new MongoClient(mongoUrl);
  await client.connect();
  const db = client.db(dbName);

  return betterAuth({
    database: mongodbAdapter(db),
    baseURL: betterAuthUrl,
    secret: process.env.BETTER_AUTH_SECRET,
    trustedOrigins,
    emailAndPassword: {
      enabled: true,
      requireEmailVerification: false,
    },
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
      // Keep cookie *name* and Secure attribute aligned. Without this, baseURL
      // https://localhost:3000 forces a `__Secure-` cookie name while
      // defaultCookieAttributes.secure is false in development — browsers then
      // reject the cookie (and Excel WebView claim would set the wrong name).
      useSecureCookies: process.env.NODE_ENV === 'production',
      defaultCookieAttributes: {
        // Same-site cookies work for same-origin Vite proxy callbacks.
        sameSite: 'lax',
        // The Server backend itself only ever serves plain HTTP in local dev
        // (Server/.env's PORT, no TLS) — HTTPS only exists on the Excel
        // add-in's own Vite dev server, which terminates TLS in front of it.
        // A Secure cookie is silently refused by the browser on any
        // non-HTTPS response, which breaks every *direct* caller of Server
        // (e.g. the Landing-page's browser tab, which talks to
        // http://localhost:4001 with no HTTPS hop in between at all) even
        // though the cookie appears to be set in the response headers.
        // betterAuthUrl.startsWith('https') worked by coincidence for the
        // Excel add-in (whose browser-facing URL genuinely is HTTPS via the
        // Vite proxy) but breaks every other direct consumer — key this off
        // NODE_ENV instead, since production Server is always behind HTTPS.
        secure: process.env.NODE_ENV === 'production',
        httpOnly: true,
        path: '/',
      },
    },
  });
}
