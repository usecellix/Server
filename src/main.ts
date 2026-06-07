import 'reflect-metadata';
import { randomUUID } from 'node:crypto';
import { ValidationPipe } from '@nestjs/common';
import { Reflector } from '@nestjs/core';
import { NestFactory } from '@nestjs/core';
import { FastifyAdapter, NestFastifyApplication } from '@nestjs/platform-fastify';
import { AppModule } from './app.module';
import { AppConfigService } from './config/app-config.service';
import { HttpExceptionFilter } from './common/filters/http-exception.filter';
import { ResponseEnvelopeInterceptor } from './common/interceptors/response-envelope.interceptor';
import { TRACE_ID_HEADER } from './common/constants/trace-id.constant';
import { setupHttpLogging } from './common/logging/http-logging';
import { getTraceIdFromHeaders } from './common/utils/trace-id.util';

async function bootstrap(): Promise<void> {
  const app = await NestFactory.create<NestFastifyApplication>(AppModule, new FastifyAdapter());
  const config = app.get(AppConfigService);
  const port = config.port;

  setupHttpLogging(app);

  app.useGlobalPipes(
    new ValidationPipe({
      whitelist: true,
      forbidNonWhitelisted: true,
      transform: true,
    }),
  );
  app.useGlobalFilters(new HttpExceptionFilter());
  app.useGlobalInterceptors(new ResponseEnvelopeInterceptor(app.get(Reflector)));
  app.enableCors();

  app
    .getHttpAdapter()
    .getInstance()
    .addHook('onRequest', (request: { headers: Record<string, unknown> }, reply: { header: (name: string, value: string) => void }, done: () => void) => {
      const traceId = getTraceIdFromHeaders(request.headers);
      request.headers[TRACE_ID_HEADER] = traceId;
      reply.header(TRACE_ID_HEADER, traceId || randomUUID());
      done();
    });

  await app.listen(port);
  console.log(`Server started on http://localhost:${port} [${config.nodeEnv}]`);
}

bootstrap();
