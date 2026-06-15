import { FastifyReply } from 'fastify';

export function initSseResponse(reply: FastifyReply): void {
  reply.raw.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
  });
}

export function writeSseEvent(reply: FastifyReply, event: string, data: unknown): void {
  reply.raw.write(`event: ${event}\n`);
  reply.raw.write(`data: ${JSON.stringify(data)}\n\n`);
  const flushable = reply.raw as { flush?: () => void };
  flushable.flush?.();
}

export function endSseResponse(reply: FastifyReply): void {
  reply.raw.end();
}
