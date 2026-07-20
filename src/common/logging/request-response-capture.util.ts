import { sanitizeLogBody } from './log-body.util';

const CAPTURE_KEY = Symbol.for('cellix.requestLogCapture');

/** Soft cap so a single log line stays readable. */
export const MAX_LOGGED_RESPONSE_CHARS = 16_384;
const MAX_SSE_EVENTS = 40;

export type CapturedResponse =
  | { kind: 'json'; body: unknown }
  | { kind: 'sse'; events: Array<{ event: string; data: unknown }> }
  | { kind: 'text'; body: string };

/** Duck type so Nest's nested fastify package and top-level `fastify` both work. */
export type ReplyWithRaw = { raw: object };

interface CaptureState {
  response?: CapturedResponse;
}

type WithCapture = {
  [CAPTURE_KEY]?: CaptureState;
};

function getState(target: object): CaptureState {
  const holder = target as WithCapture;
  if (!holder[CAPTURE_KEY]) {
    holder[CAPTURE_KEY] = {};
  }
  return holder[CAPTURE_KEY]!;
}

/** Prefer reply.raw (SSE uses Node ServerResponse); fall back to request. */
export function captureJsonResponse(reply: ReplyWithRaw, body: unknown): void {
  const state = getState(reply.raw);
  state.response = {
    kind: 'json',
    body: sanitizeLogBody(body),
  };
}

export function captureTextResponse(reply: ReplyWithRaw, body: string): void {
  const state = getState(reply.raw);
  state.response = {
    kind: 'text',
    body:
      body.length > MAX_LOGGED_RESPONSE_CHARS
        ? `${body.slice(0, MAX_LOGGED_RESPONSE_CHARS)}… [truncated]`
        : body,
  };
}

export function captureSseEvent(reply: ReplyWithRaw, event: string, data: unknown): void {
  const state = getState(reply.raw);
  if (!state.response || state.response.kind !== 'sse') {
    state.response = { kind: 'sse', events: [] };
  }
  if (state.response.events.length >= MAX_SSE_EVENTS) {
    return;
  }
  state.response.events.push({
    event,
    data: sanitizeLogBody(data),
  });
}

export function getCapturedResponse(
  reply: ReplyWithRaw,
  request?: object,
): CapturedResponse | undefined {
  const fromReply = (reply.raw as WithCapture)[CAPTURE_KEY]?.response;
  if (fromReply) return fromReply;
  if (request) {
    return (request as WithCapture)[CAPTURE_KEY]?.response;
  }
  return undefined;
}

/** Shrink oversized payloads so NDJSON lines stay bounded. */
export function summarizeResponseForLog(response: CapturedResponse | undefined): unknown {
  if (!response) return undefined;

  try {
    const json = JSON.stringify(response);
    if (json.length <= MAX_LOGGED_RESPONSE_CHARS) {
      return response;
    }
    return {
      kind: response.kind,
      truncated: true,
      preview: `${json.slice(0, MAX_LOGGED_RESPONSE_CHARS)}…`,
      ...(response.kind === 'sse'
        ? { eventCount: response.events.length, events: response.events.map((e) => e.event) }
        : {}),
    };
  } catch {
    return { kind: response.kind, error: 'unserializable' };
  }
}
