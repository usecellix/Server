import { Injectable, Logger } from '@nestjs/common';

export interface RangeDataToolRequest {
  name: 'get_range_data';
  sheet: string;
  range: string;
}

export interface RangeDataToolResult {
  values: unknown[][];
  error?: string;
}

interface PendingRequest {
  resolve: (result: RangeDataToolResult) => void;
  reject: (error: Error) => void;
  timer: ReturnType<typeof setTimeout>;
}

@Injectable()
export class ToolBridgeService {
  private readonly logger = new Logger(ToolBridgeService.name);
  private readonly pending = new Map<string, PendingRequest>();
  private readonly DEFAULT_TIMEOUT_MS = 30_000;

  async waitForRangeData(
    conversationId: string,
    request: RangeDataToolRequest,
    emit: (event: string, data: Record<string, unknown>) => void,
    timeoutMs = this.DEFAULT_TIMEOUT_MS,
  ): Promise<RangeDataToolResult> {
    const requestId = `tr_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
    const key = this.buildKey(conversationId, requestId);

    this.logger.log(
      `Tool request ${requestId}: get_range_data(${request.sheet}, ${request.range})`,
    );

    emit('tool_request', {
      requestId,
      conversationId,
      tool: request.name,
      sheet: request.sheet,
      range: request.range,
    });

    return new Promise<RangeDataToolResult>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(key);
        reject(new Error(`Tool request timed out after ${timeoutMs}ms`));
      }, timeoutMs);

      this.pending.set(key, { resolve, reject, timer });
    });
  }

  deliverResult(
    conversationId: string,
    requestId: string,
    result: RangeDataToolResult,
  ): boolean {
    const key = this.buildKey(conversationId, requestId);
    const pending = this.pending.get(key);
    if (!pending) {
      this.logger.warn(`No pending tool request for ${key}`);
      return false;
    }

    clearTimeout(pending.timer);
    this.pending.delete(key);
    pending.resolve(result);
    return true;
  }

  cancelConversation(conversationId: string): void {
    for (const [key, pending] of this.pending.entries()) {
      if (!key.startsWith(`${conversationId}:`)) continue;
      clearTimeout(pending.timer);
      pending.reject(new Error('Conversation cancelled'));
      this.pending.delete(key);
    }
  }

  private buildKey(conversationId: string, requestId: string): string {
    return `${conversationId}:${requestId}`;
  }
}
