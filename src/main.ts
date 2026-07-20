import 'reflect-metadata';
import { ValidationPipe } from '@nestjs/common';
import { Reflector } from '@nestjs/core';
import { NestFactory } from '@nestjs/core';
import { FastifyAdapter, NestFastifyApplication } from '@nestjs/platform-fastify';
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
  });
  app.useLogger(app.get(Logger));

  const config = app.get(AppConfigService);
  const port = config.port;
  const logger = app.get(Logger);
  const requestFileLogger = app.get(RequestFileLoggerService);

  const fastify = app.getHttpAdapter().getInstance();
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
<<<<<<< HEAD
  app.useGlobalInterceptors(
    new RequestResponseCaptureInterceptor(),
    new ResponseEnvelopeInterceptor(app.get(Reflector)),
  );
  app.enableCors();
=======
  app.useGlobalInterceptors(new ResponseEnvelopeInterceptor(app.get(Reflector)));
  app.enableCors({
    origin: config.clientOrigin,
    credentials: true,
  });
>>>>>>> 79b55a729d32439c8865d125c5c4c0c1a20e34a6

  await app.listen(port, '0.0.0.0');
  logger.log(`Server started on http://localhost:${port} [${config.nodeEnv}]`);
}

bootstrap();
