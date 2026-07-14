/**
 * Runtime ESM import that TypeScript will not rewrite to require().
 * Needed because better-auth is ESM-only and this Nest app compiles to CommonJS.
 */
export function importEsm<T = unknown>(specifier: string): Promise<T> {
  // eslint-disable-next-line @typescript-eslint/no-implied-eval, no-new-func
  return new Function('s', 'return import(s)')(specifier) as Promise<T>;
}
