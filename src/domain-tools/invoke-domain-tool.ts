import { DomainTool, DomainToolResult } from './types/domain-tool.types';
import { StructuredLogger } from '../agents/logging/structured-logger';

/**
 * Invoke a domain tool and emit a structured audit log with confidence/exceptions.
 * ExecutorAgent should call this rather than the registry entry directly.
 */
export function invokeDomainToolLogged<TIn, TOut>(
  toolName: string,
  tool: DomainTool<TIn, TOut>,
  input: TIn,
  structuredLogger: StructuredLogger,
  traceId: string,
): DomainToolResult<TOut> {
  const startedAt = Date.now();
  try {
    const result = tool(input);
    structuredLogger.logDomainToolCall({
      traceId,
      toolName,
      confidence: result.confidence,
      exceptionCount: result.exceptions.length,
      exceptionCodes: result.exceptions.map((e) => e.code),
      sourceRefCount: result.sourceRefs.length,
      durationMs: Date.now() - startedAt,
      success: true,
    });
    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    structuredLogger.logDomainToolCall({
      traceId,
      toolName,
      confidence: 0,
      exceptionCount: 0,
      exceptionCodes: [],
      sourceRefCount: 0,
      durationMs: Date.now() - startedAt,
      success: false,
      error: message,
    });
    throw error;
  }
}
