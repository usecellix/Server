import { DomainTool, DomainToolResult } from '../types/domain-tool.types';

/** Compile-time + runtime guard: DomainToolResult must always carry confidence/exceptions. */
export function assertDomainToolResultShape<T>(result: DomainToolResult<T>): void {
  expect(typeof result.confidence).toBe('number');
  expect(Array.isArray(result.exceptions)).toBe(true);
  expect(Array.isArray(result.sourceRefs)).toBe(true);
  expect(result.data).toBeDefined();
}

export function expectStubNotImplemented<TIn, TOut>(
  tool: DomainTool<TIn, TOut>,
  input: TIn,
): void {
  expect(typeof tool).toBe('function');
  expect(() => tool(input)).toThrow(/Not implemented/i);
}
