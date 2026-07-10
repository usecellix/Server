declare module '@toon-format-cjs/toon' {
  export function encode(input: unknown, options?: Record<string, unknown>): string;
  export function decode(input: string, options?: Record<string, unknown>): unknown;
}
