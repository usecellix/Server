import { CallHandler, ExecutionContext, Injectable, NestInterceptor } from '@nestjs/common';
import { Reflector } from '@nestjs/core';
import { Observable } from 'rxjs';
import { map } from 'rxjs/operators';
import { FastifyRequest } from 'fastify';
import { ApiSuccessResponse } from '../interfaces/api-response.interface';
import { TRACE_ID_HEADER } from '../constants/trace-id.constant';
import { normalizeTraceId } from '../utils/trace-id.util';
import { SKIP_ENVELOPE_KEY } from '../decorators/skip-envelope.decorator';

@Injectable()
export class ResponseEnvelopeInterceptor<T> implements NestInterceptor<T, unknown> {
  constructor(private readonly reflector: Reflector) {}

  intercept(context: ExecutionContext, next: CallHandler<T>): Observable<unknown> {
    const skipEnvelope = this.reflector.getAllAndOverride<boolean>(SKIP_ENVELOPE_KEY, [
      context.getHandler(),
      context.getClass(),
    ]);

    if (skipEnvelope) {
      return next.handle();
    }

    const request = context.switchToHttp().getRequest<FastifyRequest>();
    const traceId = normalizeTraceId(request.headers[TRACE_ID_HEADER]);

    return next.handle().pipe(
      map((data) => ({
        success: true,
        traceId,
        data,
      })),
    );
  }
}
