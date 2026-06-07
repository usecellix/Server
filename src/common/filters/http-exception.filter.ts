import { ArgumentsHost, Catch, ExceptionFilter, HttpException, HttpStatus } from '@nestjs/common';
import { FastifyReply, FastifyRequest } from 'fastify';
import { ApiErrorResponse } from '../interfaces/api-response.interface';
import { TRACE_ID_HEADER } from '../constants/trace-id.constant';
import { normalizeTraceId } from '../utils/trace-id.util';

@Catch()
export class HttpExceptionFilter implements ExceptionFilter {
  catch(exception: unknown, host: ArgumentsHost): void {
    const ctx = host.switchToHttp();
    const request = ctx.getRequest<FastifyRequest>();
    const response = ctx.getResponse<FastifyReply>();
    const traceId = normalizeTraceId(request.headers[TRACE_ID_HEADER]);

    const status = exception instanceof HttpException ? exception.getStatus() : HttpStatus.INTERNAL_SERVER_ERROR;
    const payload = exception instanceof HttpException ? exception.getResponse() : null;

    const errorResponse: ApiErrorResponse = {
      success: false,
      traceId,
      error: {
        code: this.getErrorCode(status),
        message: this.getErrorMessage(payload),
        details: this.getErrorDetails(payload),
      },
    };

    response.status(status).send(errorResponse);
  }

  private getErrorCode(status: number): string {
    switch (status) {
      case HttpStatus.BAD_REQUEST:
        return 'BAD_REQUEST';
      case HttpStatus.UNAUTHORIZED:
        return 'UNAUTHORIZED';
      case HttpStatus.FORBIDDEN:
        return 'FORBIDDEN';
      case HttpStatus.NOT_FOUND:
        return 'NOT_FOUND';
      case HttpStatus.CONFLICT:
        return 'CONFLICT';
      case HttpStatus.UNPROCESSABLE_ENTITY:
        return 'VALIDATION_ERROR';
      default:
        return status >= 500 ? 'INTERNAL_SERVER_ERROR' : 'HTTP_ERROR';
    }
  }

  private getErrorMessage(payload: unknown): string {
    if (typeof payload === 'string') {
      return payload;
    }

    if (typeof payload === 'object' && payload !== null && 'message' in payload) {
      const message = (payload as { message: unknown }).message;
      return Array.isArray(message) ? message.join(', ') : String(message);
    }

    return 'Unexpected error occurred';
  }

  private getErrorDetails(payload: unknown): unknown {
    if (typeof payload === 'object' && payload !== null) {
      return payload;
    }
    return undefined;
  }
}
