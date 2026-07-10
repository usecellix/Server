import { randomUUID } from 'node:crypto';
import { Params } from 'nestjs-pino';
import { TRACE_ID_HEADER } from '../constants/trace-id.constant';
import { sanitizeLogBody } from './log-body.util';

function readTraceId(headers: Record<string, unknown>): string {
  const value = headers[TRACE_ID_HEADER];
  if (typeof value === 'string' && value.trim()) {
    return value.trim();
  }
  if (Array.isArray(value) && typeof value[0] === 'string' && value[0].trim()) {
    return value[0].trim();
  }
  return randomUUID();
}

export function buildPinoParams(nodeEnv: string): Params {
  const isProduction = nodeEnv === 'production';

  return {
    pinoHttp: {
      level: isProduction ? 'info' : 'debug',
      genReqId: (req, res) => {
        const traceId = readTraceId(req.headers as Record<string, unknown>);
        req.headers[TRACE_ID_HEADER] = traceId;
        res.setHeader(TRACE_ID_HEADER, traceId);
        return traceId;
      },
      customProps: (req) => ({
        traceId: readTraceId(req.headers as Record<string, unknown>),
      }),
      customLogLevel: (_req, res, err) => {
        if (err || res.statusCode >= 500) return 'error';
        if (res.statusCode >= 400) return 'warn';
        return 'info';
      },
      serializers: {
        req(req) {
          const rawBody =
            (req as { raw?: { body?: unknown } }).raw?.body ??
            (req as { body?: unknown }).body;
          return {
            id: req.id,
            method: req.method,
            url: req.url,
            traceId: readTraceId(req.headers as Record<string, unknown>),
            body: sanitizeLogBody(rawBody),
          };
        },
        res(res) {
          const contentType = res.getHeader?.('content-type');
          const isSse =
            typeof contentType === 'string' && contentType.includes('text/event-stream');
          return {
            statusCode: res.statusCode,
            ...(isSse ? { body: '[SSE stream omitted]' } : {}),
          };
        },
      },
      redact: {
        paths: ['req.headers.authorization', 'req.headers.cookie'],
        remove: true,
      },
    },
  };
}
