import { Injectable, Logger, OnModuleInit } from '@nestjs/common';
import { HttpAdapterHost } from '@nestjs/core';
import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import { getAuth } from './auth';

/**
 * Mounts Better Auth's catch-all handler on Fastify at /api/auth/*.
 */
@Injectable()
export class AuthRouteRegistrar implements OnModuleInit {
  private readonly logger = new Logger(AuthRouteRegistrar.name);

  constructor(private readonly httpAdapterHost: HttpAdapterHost) {}

  async onModuleInit(): Promise<void> {
    const auth = await getAuth();
    const instance = this.httpAdapterHost.httpAdapter.getInstance() as FastifyInstance;

    instance.route({
      method: ['GET', 'POST'],
      url: '/api/auth/*',
      config: {
        rawBody: true,
      },
      handler: async (request: FastifyRequest, reply: FastifyReply) => {
        // Prefer proxy headers from Vite HTTPS; fall back to BETTER_AUTH_URL scheme.
        const configured = process.env.BETTER_AUTH_URL || 'https://localhost:3000';
        let configuredOrigin: URL | null = null;
        try {
          configuredOrigin = new URL(configured);
        } catch {
          configuredOrigin = null;
        }

        const host =
          (request.headers['x-forwarded-host'] as string) ||
          request.headers.host ||
          configuredOrigin?.host ||
          'localhost:3000';
        const protocol =
          (request.headers['x-forwarded-proto'] as string) ||
          configuredOrigin?.protocol.replace(':', '') ||
          'https';
        const url = new URL(request.url, `${protocol}://${host}`);

        const headers = new Headers();
        for (const [key, value] of Object.entries(request.headers)) {
          if (value === undefined) continue;
          if (Array.isArray(value)) {
            for (const entry of value) headers.append(key, entry);
          } else {
            headers.set(key, value);
          }
        }
        // Ensure Better Auth sees the public HTTPS origin behind the Vite proxy.
        if (!headers.has('x-forwarded-proto')) headers.set('x-forwarded-proto', protocol);
        if (!headers.has('x-forwarded-host')) headers.set('x-forwarded-host', host);

        const hasBody = request.method !== 'GET' && request.method !== 'HEAD';
        const webRequest = new Request(url.toString(), {
          method: request.method,
          headers,
          body: hasBody ? await readRawBody(request) : undefined,
          // @ts-expect-error duplex required for streaming request bodies in Node fetch
          duplex: hasBody ? 'half' : undefined,
        });

        const response = await auth.handler(webRequest);

        reply.status(response.status);
        response.headers.forEach((value, key) => {
          // Fastify manages content-length itself when sending a body.
          if (key.toLowerCase() === 'content-length') return;
          reply.header(key, value);
        });

        const buffer = Buffer.from(await response.arrayBuffer());
        return reply.send(buffer.length ? buffer : null);
      },
    });

    this.logger.log('Better Auth routes mounted at /api/auth/*');
  }
}

async function readRawBody(request: FastifyRequest): Promise<string | undefined> {
  const raw = (request as FastifyRequest & { rawBody?: Buffer | string }).rawBody;
  if (typeof raw === 'string') return raw;
  if (Buffer.isBuffer(raw)) return raw.toString('utf8');

  if (request.body === undefined || request.body === null) {
    return undefined;
  }

  if (typeof request.body === 'string') {
    return request.body;
  }

  return JSON.stringify(request.body);
}
