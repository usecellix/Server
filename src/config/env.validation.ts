import * as Joi from 'joi';

export const envValidationSchema = Joi.object({
  NODE_ENV: Joi.string().valid('development', 'test', 'production').default('development'),
  PORT: Joi.number().port().default(4001),
  MONGODB_URL: Joi.string().uri({ scheme: [/mongodb(\+srv)?/] }).default('mongodb://127.0.0.1:27017/cellix'),
  MONGODB_DB_NAME: Joi.string().trim().min(1).default('cellix'),
  OPENROUTER_API_KEY: Joi.string().allow('').optional(),
  OPENROUTER_MODEL: Joi.string().optional(),
  OPENROUTER_MODEL_LOW: Joi.string().default('openai/gpt-5-mini'),
  OPENROUTER_MODEL_MEDIUM: Joi.string().default('openai/gpt-5-mini'),
  OPENROUTER_MODEL_HIGH: Joi.string().default('openai/gpt-5'),
  OPENROUTER_MODEL_TIER1: Joi.string().optional(),
  OPENROUTER_MODEL_PLANNER: Joi.string().optional(),
  // LlmRouterService's LOW-tier classifier call — decoupled from the shared
  // LOW tier (which multi-sheet summaries and ambiguity clarification also
  // use) so a router-specific model eval doesn't affect those. Defaults to
  // openRouterModelLow, so unset is a no-op.
  OPENROUTER_MODEL_ROUTER: Joi.string().optional(),
  // Tier2GenerateVerifyService's ExecutorAgent.execute() call — decoupled from
  // Tier 3's Executor (same ExecutorAgent class, same openRouterModelHigh by
  // default) so a Tier-2-only model eval doesn't touch Tier 3. Unset is a
  // no-op (falls back to openRouterModelHigh, today's shared behavior).
  OPENROUTER_MODEL_TIER2_GENERATE: Joi.string().optional(),
  // MODEL_PROFILE=dev swaps every tier's resolved model string to
  // CELLIX_DEV_MODEL, independent of tier-routing logic (AppConfigService is
  // the sole chokepoint every tier already reads its model through).
  MODEL_PROFILE: Joi.string().valid('prod', 'dev').default('prod'),
  CELLIX_DEV_MODEL: Joi.string().default('z-ai/glm-5.3-flash'),
  OPENROUTER_HTTP_REFERER: Joi.string().default('https://cellix.local'),
  OPENAI_API_KEY: Joi.string().allow('').optional(),
  OPENAI_MODEL: Joi.string().default('gpt-4o-mini'),
  OPENAI_MODEL_LOW: Joi.string().optional(),
  OPENAI_MODEL_MEDIUM: Joi.string().optional(),
  OPENAI_MODEL_HIGH: Joi.string().optional(),
  ENABLE_COMPLEXITY_TIERING: Joi.string()
    .valid('off', 'shadow', 'tier01', 'tier0-1', 'tier0_1', 'full', 'on', 'true', 'false', '0', '1')
    .optional(),
  /**
   * Step-wise Tier 3 execution (TASKS.md #153, STEPWISE_EXECUTION.md). Defaults
   * off everywhere — it changes the SSE contract (a run ends with `wave_ready`,
   * not `conversation_end`), so an add-in build predating the client half must
   * not start receiving paused runs it will never continue.
   */
  ENABLE_STEPWISE_EXECUTION: Joi.string()
    .valid('on', 'off', 'true', 'false', '0', '1')
    .optional(),
  /** Retention for `agent_runs` working state — see agent-run.schema.ts. */
  AGENT_RUN_TTL_HOURS: Joi.number().positive().optional(),
  // Lets eval/run-live-eval.ts authenticate against a real running backend
  // without a browser session. AuthGuard only honors this when NODE_ENV is
  // NOT 'production' AND this is set — unset (the default) changes nothing
  // in prod. See auth.guard.ts.
  CELLIX_EVAL_BYPASS_TOKEN: Joi.string().optional(),
  // Better Auth / OAuth — required at runtime for social login; optional so the API can boot without them.
  BETTER_AUTH_SECRET: Joi.string().min(32).optional(),
  BETTER_AUTH_URL: Joi.string().uri().optional(),
  CLIENT_ORIGIN: Joi.string().uri().default('https://localhost:3000'),
  GOOGLE_CLIENT_ID: Joi.string().allow('').optional(),
  GOOGLE_CLIENT_SECRET: Joi.string().allow('').optional(),
  MICROSOFT_CLIENT_ID: Joi.string().allow('').optional(),
  MICROSOFT_CLIENT_SECRET: Joi.string().allow('').optional(),
  MICROSOFT_TENANT_ID: Joi.string().default('common'),
  // Razorpay — CREDIT_SYSTEM.md §7 (migrated from Stripe, credit-system-v2
  // session — the original pricing doc always specified Razorpay). Optional
  // so the API can boot without them (checkout creation/webhook processing
  // fail cleanly instead, same pattern as OPENROUTER_API_KEY).
  // RAZORPAY_WEBHOOK_SECRET verifies webhook signatures; without it,
  // incoming webhooks are rejected rather than trusted unverified.
  // RAZORPAY_PLAN_ID_* reference Plans created ahead of time via Razorpay's
  // dashboard/API (a one-time setup step outside this codebase, same
  // constraint Stripe Price IDs had) — Solo/Firm/Beta each need their own.
  RAZORPAY_KEY_ID: Joi.string().allow('').optional(),
  RAZORPAY_KEY_SECRET: Joi.string().allow('').optional(),
  RAZORPAY_WEBHOOK_SECRET: Joi.string().allow('').optional(),
  RAZORPAY_PLAN_ID_SOLO: Joi.string().allow('').optional(),
  RAZORPAY_PLAN_ID_FIRM: Joi.string().allow('').optional(),
  RAZORPAY_PLAN_ID_BETA: Joi.string().allow('').optional(),
  // Where Razorpay's hosted checkout redirects after payment — the marketing
  // site, not this API. Defaults to the local landing-page dev port.
  CHECKOUT_SUCCESS_URL: Joi.string().uri().default('http://localhost:5173/checkout/success'),
  CHECKOUT_CANCEL_URL: Joi.string().uri().default('http://localhost:5173/checkout'),
});
