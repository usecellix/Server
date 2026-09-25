import 'reflect-metadata';
import { ValidationPipe } from '@nestjs/common';
import { Reflector } from '@nestjs/core';
import { NestFactory } from '@nestjs/core';
import { FastifyAdapter, NestFastifyApplication } from '@nestjs/platform-fastify';
import fastifyHelmet from '@fastify/helmet';
import fastifyRateLimit from '@fastify/rate-limit';
import { Logger } from 'nestjs-pino';
import { AppModule } from './app.module';
import { AppConfigService } from './config/app-config.service';
import { TRACE_ID_HEADER } from './common/constants/trace-id.constant';
import { HttpExceptionFilter } from './common/filters/http-exception.filter';
import { RequestResponseCaptureInterceptor } from './common/interceptors/request-response-capture.interceptor';
import { ResponseEnvelopeInterceptor } from './common/interceptors/response-envelope.interceptor';
import { RequestFileLoggerService } from './common/logging/request-file-logger.service';
import {
  getCapturedResponse,
  summarizeResponseForLog,
} from './common/logging/request-response-capture.util';

function clipMessage(value: unknown, max = 200): string | undefined {
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  return trimmed.length <= max ? trimmed : `${trimmed.slice(0, max)}…`;
}

/** Undici/OpenRouter often emits a second rejection after ECONNRESET — do not crash nodemon. */
function isBenignNetworkAbort(error: unknown): boolean {
  if (!error || typeof error !== 'object') return false;
  const err = error as Error & { cause?: { code?: string }; code?: string };
  const message = String(err.message ?? '').toLowerCase();
  const code = err.code ?? err.cause?.code;
  return (
    message === 'terminated' ||
    message.includes('econnreset') ||
    message.includes('fetch failed') ||
    code === 'ECONNRESET' ||
    code === 'ECONNREFUSED' ||
    code === 'ETIMEDOUT' ||
    code === 'UND_ERR_SOCKET'
  );
}

function installProcessSafetyNets(): void {
  process.on('unhandledRejection', (reason) => {
    if (isBenignNetworkAbort(reason)) {
      // eslint-disable-next-line no-console
      console.warn(
        '[Cellix] Swallowed unhandled network rejection (LLM connection reset):',
        reason instanceof Error ? reason.message : reason,
      );
      return;
    }
    // eslint-disable-next-line no-console
    console.error('[Cellix] Unhandled promise rejection:', reason);
  });

  process.on('uncaughtException', (error) => {
    if (isBenignNetworkAbort(error)) {
      // eslint-disable-next-line no-console
      console.warn(
        '[Cellix] Swallowed uncaught network exception (LLM connection reset):',
        error.message,
      );
      return;
    }
    // eslint-disable-next-line no-console
    console.error('[Cellix] Uncaught exception:', error);
    process.exit(1);
  });
}

async function bootstrap(): Promise<void> {
  installProcessSafetyNets();

  const app = await NestFactory.create<NestFastifyApplication>(AppModule, new FastifyAdapter(), {
    bufferLogs: true,
    // Razorpay webhook signature verification needs the exact request bytes
    // on req.rawBody — Nest registers the JSON parser once with this enabled.
    rawBody: true,
  });
  app.useLogger(app.get(Logger));

  const config = app.get(AppConfigService);
  const port = config.port;
  const logger = app.get(Logger);
  const requestFileLogger = app.get(RequestFileLoggerService);

  const fastify = app.getHttpAdapter().getInstance();

  // `npm ls fastify` shows two distinct installed versions (5.8.4 nested
  // under @nestjs/platform-fastify, 5.8.5 at top level, both real Fastify 5 —
  // not a genuine incompatibility). Each plugin's .d.ts resolves `fastify`'s
  // types against whichever copy its own node_modules sees, so
  // NestFastifyApplication's FastifyInstance and each plugin's expected
  // FastifyInstance are structurally close but nominally different types to
  // tsc, even though they're the same object at runtime. `pluginFastify`
  // narrows the cast to these two registrations only, rather than casting
  // the shared `fastify` binding everywhere it's used below.
  const pluginFastify = fastify as unknown as Parameters<typeof fastifyHelmet>[0] &
    Parameters<typeof fastifyRateLimit>[0];

  // TASKS.md #344 — this is a pure JSON/SSE API, never serves HTML, so a CSP
  // (helmet's default) has nothing to constrain and only risks an unexpected
  // interaction with the SSE response's manually-written headers
  // (initSseResponse writes them straight to the raw Node response, bypassing
  // Fastify's onSend hooks entirely — helmet cannot touch those either way,
  // but leaving CSP off keeps the header set honest about what this API is).
  // The other defaults (X-Content-Type-Options, X-Frame-Options, etc.) are
  // still useful even for a JSON API against MIME-sniffing/embedding.
  await pluginFastify.register(fastifyHelmet, { contentSecurityPolicy: false });

  // Coarse, IP-keyed defense-in-depth against anonymous/pre-auth abuse
  // (unauthenticated endpoints, credential stuffing on /auth/*, scripted
  // scanning). Generous on purpose — the real per-user protection against
  // unlimited concurrent LLM spend is ConcurrencyLimitService, applied only
  // to the routes that actually spend credit and keyed by user, not IP,
  // which this plugin cannot see until Nest's AuthGuard has run.
  await pluginFastify.register(fastifyRateLimit, {
    max: 300,
    timeWindow: '1 minute',
  });

  fastify.addHook('onResponse', (request, reply, done) => {
    const body = (request as { body?: { message?: unknown } }).body;
    const headerTrace = request.headers[TRACE_ID_HEADER];
    const traceId =
      typeof headerTrace === 'string'
        ? headerTrace
        : Array.isArray(headerTrace)
          ? headerTrace[0]
          : undefined;

    const response = summarizeResponseForLog(getCapturedResponse(reply, request));

    requestFileLogger.logRequest({
      method: request.method,
      url: request.url,
      statusCode: reply.statusCode,
      responseTimeMs: Math.round(reply.elapsedTime ?? 0),
      reqId: typeof request.id === 'string' ? request.id : String(request.id ?? ''),
      traceId,
      message: clipMessage(body?.message),
      ...(response !== undefined ? { response } : {}),
    });
    done();
  });

  app.useGlobalPipes(
    new ValidationPipe({
      whitelist: true,
      forbidNonWhitelisted: true,
      transform: true,
    }),
  );
  app.useGlobalFilters(new HttpExceptionFilter());
  app.useGlobalInterceptors(
    new RequestResponseCaptureInterceptor(),
    new ResponseEnvelopeInterceptor(app.get(Reflector)),
  );
  app.enableCors({
    origin: config.allowedCorsOrigins,
    credentials: true,
  });

  await app.listen(port, '0.0.0.0');
  logger.log(`Server started on http://localhost:${port} [${config.nodeEnv}]`);
}

bootstrap();
