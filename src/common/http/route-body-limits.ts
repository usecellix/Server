/**
 * Per-route request size limits.
 *
 * Fastify refuses a body over 1 MiB by default, which is right for almost every
 * route here. The few that legitimately take more are listed in this file, so
 * the default is never raised for the whole API to suit one of them.
 */
export const ROUTE_BODY_LIMITS: Record<string, number> = {
  // A decoded bank statement: six months measures about 0.4 MB, so this covers many years.
  '/ingest/bank-statement': 16 * 1024 * 1024,
};

interface RouteOptionsLike {
  url: string;
  bodyLimit?: number;
}

interface FastifyLike {
  addHook(name: 'onRoute', hook: (routeOptions: RouteOptionsLike) => void): unknown;
}

/**
 * Must be called before the routes are registered (before `app.listen` /
 * `app.init`): an `onRoute` hook only sees routes added after it.
 */
export function applyRouteBodyLimits(
  fastify: FastifyLike,
  limits: Record<string, number> = ROUTE_BODY_LIMITS,
): void {
  fastify.addHook('onRoute', (routeOptions) => {
    const limit = limits[routeOptions.url];
    if (limit !== undefined) routeOptions.bodyLimit = limit;
  });
}
