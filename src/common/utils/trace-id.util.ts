import { randomUUID } from 'node:crypto';
import { TRACE_ID_HEADER } from '../constants/trace-id.constant';

export function normalizeTraceId(value: unknown): string {
  return typeof value === 'string' && value.trim().length > 0 ? value : randomUUID();
}

export function getTraceIdFromHeaders(headers: Record<string, unknown>): string {
  return normalizeTraceId(headers[TRACE_ID_HEADER]);
}
