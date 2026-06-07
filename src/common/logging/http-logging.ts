import type { NestFastifyApplication } from '@nestjs/platform-fastify';
import type { IncomingMessage, ServerResponse } from 'node:http';
import morgan from 'morgan';
import { TRACE_ID_HEADER } from '../constants/trace-id.constant';

const MAX_LOG_LENGTH = 4096;

type RawWithBody = IncomingMessage & { body?: unknown };
type ResponseWithPayload = ServerResponse & { _logPayload?: unknown };
type JsonRecord = Record<string, unknown>;

function truncate(value: string): string {
  if (value.length <= MAX_LOG_LENGTH) {
    return value;
  }
  return `${value.slice(0, MAX_LOG_LENGTH)}… [truncated ${value.length - MAX_LOG_LENGTH} chars]`;
}

function serializeBody(body: unknown): string {
  if (body === undefined || body === null) {
    return '-';
  }
  const sanitized = sanitizeBody(body);
  if (typeof body === 'string') {
    return truncate(body);
  }
  if (Buffer.isBuffer(body)) {
    return truncate(body.toString('utf8'));
  }
  try {
    return truncate(JSON.stringify(sanitized));
  } catch {
    return '[unserializable]';
  }
}

function isRecord(value: unknown): value is JsonRecord {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function summarizeSheetData(sheetData: unknown): JsonRecord | string {
  if (!Array.isArray(sheetData)) {
    return '[invalid sheetData]';
  }

  const firstRow = sheetData[0];
  const columnCount = Array.isArray(firstRow) ? firstRow.length : 0;
  const headers = Array.isArray(firstRow) ? firstRow.slice(0, 12) : [];

  return {
    rows: sheetData.length,
    columns: columnCount,
    headers,
  };
}

function sanitizeBody(body: unknown): unknown {
  if (!isRecord(body)) {
    return body;
  }

  const sanitized: JsonRecord = { ...body };

  if ('sheetData' in sanitized) {
    sanitized.sheetData = summarizeSheetData(sanitized.sheetData);
  }

  if (isRecord(sanitized.context) && Array.isArray(sanitized.context.previousMessages)) {
    sanitized.context = {
      ...sanitized.context,
      previousMessages: `${sanitized.context.previousMessages.length} message(s)`,
    };
  }

  return sanitized;
}

function isSseResponse(res: ServerResponse): boolean {
  const contentType = res.getHeader('content-type');
  return typeof contentType === 'string' && contentType.includes('text/event-stream');
}

export function setupHttpLogging(app: NestFastifyApplication): void {
  const fastify = app.getHttpAdapter().getInstance();

  fastify.addHook('preHandler', async (request, _reply) => {
    (request.raw as RawWithBody).body = request.body;
  });

  fastify.addHook('onSend', async (_request, reply, payload) => {
    (reply.raw as ResponseWithPayload)._logPayload = payload;
    return payload;
  });

  morgan.token('trace-id', (req) => {
    const traceId = req.headers[TRACE_ID_HEADER];
    return typeof traceId === 'string' ? traceId : '-';
  });

  morgan.token('req-body', (req) => serializeBody((req as RawWithBody).body));

  morgan.token('res-body', (_req, res) => {
    const response = res as ResponseWithPayload;
    if (isSseResponse(response)) {
      return '[SSE stream omitted]';
    }
    return serializeBody(response._logPayload);
  });

  const format =
    ':date[iso] trace=:trace-id :method :url :status :res[content-length] :response-time ms | req=:req-body | res=:res-body';

  // Nest FastifyAdapter registers @fastify/middie during init(); only queue middleware here.
  app.use(morgan(format));
}
