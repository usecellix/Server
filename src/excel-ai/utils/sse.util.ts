import { FastifyReply } from 'fastify';
import { captureSseEvent } from '../../common/logging/request-response-capture.util';

export function initSseResponse(reply: FastifyReply): void {
  reply.raw.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
  });
}

export function writeSseEvent(reply: FastifyReply, event: string, data: unknown): void {
  captureSseEvent(reply, event, data);
  reply.raw.write(`event: ${event}\n`);
  reply.raw.write(`data: ${JSON.stringify(data)}\n\n`);
  const flushable = reply.raw as { flush?: () => void };
  flushable.flush?.();
}

export function endSseResponse(reply: FastifyReply): void {
  reply.raw.end();
}

/**
 * An AbortSignal that fires when the underlying HTTP connection for this SSE
 * response closes — the client navigating away, closing the taskpane, or the
 * "Stop" button aborting its fetch. Without this, the agentic loop had no way
 * to learn a run was cancelled: it kept executing every remaining wave (LLM
 * calls included) to completion, only to write to a response nobody was
 * reading. Call once per request, right after `initSseResponse`, and pass the
 * `signal` down into anything long-running (AgenticLoopService's wave loop).
 * TASKS.md #260.
 */
export function createRequestAbortSignal(reply: FastifyReply): AbortSignal {
  const controller = new AbortController();
  const onClose = () => controller.abort();
  reply.raw.once('close', onClose);
  reply.raw.once('error', onClose);
  return controller.signal;
}
