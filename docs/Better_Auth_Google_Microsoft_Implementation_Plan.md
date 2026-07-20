# Implementation Plan: Google & Microsoft Authentication with Better Auth

## Stack

- **Backend:** Node.js + NestJS
- **Frontend:** React (Vite recommended)
- **Authentication:** Better Auth
- **Providers:** Google OAuth + Microsoft Entra ID (Generic OAuth)

---

# 1. Project Overview

This document describes the implementation plan for integrating **Google** and **Microsoft (Entra ID)** authentication into a **NestJS** backend and **React** frontend using **Better Auth**.

## Goals

- Google authentication
- Microsoft authentication
- Cookie-based session management
- Protected backend APIs
- React authentication hooks
- Account linking support
- Production-ready configuration

---

# 2. Prerequisites

- NestJS v10+
- React (Vite recommended)
- PostgreSQL/MySQL
- Drizzle ORM or Prisma
- Google Cloud OAuth credentials
- Microsoft Entra ID App Registration
- Environment variables configured

---

# 3. Backend Implementation (NestJS)

## 3.1 Install Dependencies

```bash
npm install better-auth @thallesp/nestjs-better-auth
npm install @better-auth/drizzle-adapter
```

> Replace the database adapter if you are using Prisma or another supported adapter.

---

## 3.2 Create Auth Configuration

**File:** `src/auth.ts`

```ts
import { betterAuth } from "better-auth";
import { drizzleAdapter } from "@better-auth/drizzle-adapter";
import { genericOAuth } from "better-auth/plugins";
import { db } from "./db";

export const auth = betterAuth({
  database: drizzleAdapter(db, {
    provider: "pg",
  }),

  socialProviders: {
    google: {
      clientId: process.env.GOOGLE_CLIENT_ID!,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET!,
    },
  },

  plugins: [
    genericOAuth({
      config: [
        {
          providerId: "microsoft",
          clientId: process.env.MICROSOFT_CLIENT_ID!,
          clientSecret: process.env.MICROSOFT_CLIENT_SECRET!,
          tenantId: process.env.MICROSOFT_TENANT_ID || "common",
          scopes: [
            "openid",
            "email",
            "profile",
            "User.Read",
          ],
        },
      ],
    }),
  ],

  baseURL: process.env.BETTER_AUTH_URL || "http://localhost:3001",

  trustedOrigins: [
    "http://localhost:3000",
  ],
});
```

---

## 3.3 Environment Variables

```env
BETTER_AUTH_URL=http://localhost:3001

GOOGLE_CLIENT_ID=your_google_client_id
GOOGLE_CLIENT_SECRET=your_google_client_secret

MICROSOFT_CLIENT_ID=your_ms_client_id
MICROSOFT_CLIENT_SECRET=your_ms_client_secret
MICROSOFT_TENANT_ID=common
```

---

## 3.4 Integrate Better Auth with NestJS

### main.ts

```ts
import { NestFactory } from "@nestjs/core";
import { AppModule } from "./app.module";

async function bootstrap() {
  const app = await NestFactory.create(AppModule, {
    bodyParser: false,
  });

  app.enableCors({
    origin: "http://localhost:3000",
    credentials: true,
  });

  await app.listen(3001);
}

bootstrap();
```

### app.module.ts

```ts
import { Module } from "@nestjs/common";
import { AuthModule } from "@thallesp/nestjs-better-auth";
import { auth } from "./auth";

@Module({
  imports: [
    AuthModule.forRoot({
      auth,
    }),
  ],
})
export class AppModule {}
```

---

## 3.5 Protected Routes

```ts
import { Controller, Get } from "@nestjs/common";
import {
  Session,
  UserSession,
  AllowAnonymous,
} from "@thallesp/nestjs-better-auth";

@Controller("auth")
export class AuthController {

  @Get("me")
  getProfile(@Session() session: UserSession) {
    return {
      user: session?.user,
    };
  }

  @Get("public")
  @AllowAnonymous()
  publicRoute() {
    return {
      message: "Public endpoint",
    };
  }
}
```

---

# 4. Frontend Implementation (React)

## 4.1 Install Better Auth Client

```bash
npm install better-auth
```

---

## 4.2 Create Auth Client

**File:** `src/lib/auth-client.ts`

```ts
import { createAuthClient } from "better-auth/client";
import { genericOAuthClient } from "better-auth/client/plugins";

export const authClient = createAuthClient({
  baseURL: "http://localhost:3001",

  plugins: [
    genericOAuthClient(),
  ],
});
```

---

## 4.3 Login Component

```tsx
import { authClient } from "@/lib/auth-client";

export default function Login() {

  const signInGoogle = () => {
    authClient.signIn.social({
      provider: "google",
      callbackURL: "/dashboard",
    });
  };

  const signInMicrosoft = () => {
    authClient.signIn.oauth2({
      providerId: "microsoft",
      callbackURL: "/dashboard",
    });
  };

  return (
    <div>
      <button onClick={signInGoogle}>
        Sign in with Google
      </button>

      <button onClick={signInMicrosoft}>
        Sign in with Microsoft
      </button>
    </div>
  );
}
```

---

## 4.4 Session Management

```tsx
import { useSession } from "better-auth/react";

export default function Dashboard() {
  const {
    data: session,
    isPending,
  } = useSession();

  if (isPending) {
    return <p>Loading...</p>;
  }

  if (!session) {
    return <p>Not authenticated</p>;
  }

  return (
    <div>
      <h1>Welcome {session.user.name}</h1>

      <p>{session.user.email}</p>

      <pre>
        {JSON.stringify(session, null, 2)}
      </pre>
    </div>
  );
}
```

---

# 5. OAuth Provider Configuration

## Google

- Create OAuth Client in Google Cloud Console
- Redirect URI

```
http://localhost:3001/api/auth/callback/google
```

Scopes

- openid
- email
- profile

---

## Microsoft (Entra ID)

Register an application inside Azure Portal.

Redirect URI

```
http://localhost:3001/api/auth/oauth2/callback/microsoft
```

Permissions

- openid
- email
- profile
- User.Read

Recommended tenant

```
common
```

---

# 6. Testing Checklist

- Start backend
- Start frontend
- Test Google login
- Test Microsoft login
- Verify cookies
- Verify session persistence
- Verify protected APIs
- Test logout
- Test account linking
- Update production URLs
- Enable secure cookies

---

# 7. Optional Features

- Account linking
- Additional OAuth scopes
- Role-based authorization
- Email/password authentication
- Custom user mapping
- Multi-tenant support

Example:

```ts
authClient.linkSocial({
  provider: "google",
});
```

---

# 8. Troubleshooting

## CORS

- Check `trustedOrigins`
- Enable `credentials: true`

## Callback Errors

- Verify redirect URLs
- Verify OAuth credentials

## Body Parser Error

Ensure:

```ts
bodyParser: false
```

## Microsoft Login Issues

Try:

```ts
tenantId: "common"
```

---

# Production Checklist

- Update callback URLs
- Update trusted origins
- Enable HTTPS
- Enable secure cookies
- Store secrets securely
- Verify OAuth redirect URIs
- Test production login flow

---

# References

- Better Auth Documentation: https://better-auth.com

---

## Summary

This implementation provides:

- Google authentication
- Microsoft authentication
- Cookie-based session management
- Protected APIs
- React authentication hooks
- Better Auth integration with NestJS
- Production-ready configuration
