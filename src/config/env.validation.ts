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
  OPENROUTER_HTTP_REFERER: Joi.string().default('https://cellix.local'),
  OPENAI_API_KEY: Joi.string().allow('').optional(),
  OPENAI_MODEL: Joi.string().default('gpt-4o-mini'),
  OPENAI_MODEL_LOW: Joi.string().optional(),
  OPENAI_MODEL_MEDIUM: Joi.string().optional(),
  OPENAI_MODEL_HIGH: Joi.string().optional(),
<<<<<<< HEAD
  ENABLE_COMPLEXITY_TIERING: Joi.string()
    .valid('off', 'shadow', 'tier01', 'tier0-1', 'tier0_1', 'full', 'on', 'true', 'false', '0', '1')
    .optional(),
=======
  // Better Auth / OAuth — required at runtime for social login; optional so the API can boot without them.
  BETTER_AUTH_SECRET: Joi.string().min(32).optional(),
  BETTER_AUTH_URL: Joi.string().uri().optional(),
  CLIENT_ORIGIN: Joi.string().uri().default('https://localhost:3000'),
  GOOGLE_CLIENT_ID: Joi.string().allow('').optional(),
  GOOGLE_CLIENT_SECRET: Joi.string().allow('').optional(),
  MICROSOFT_CLIENT_ID: Joi.string().allow('').optional(),
  MICROSOFT_CLIENT_SECRET: Joi.string().allow('').optional(),
  MICROSOFT_TENANT_ID: Joi.string().default('common'),
>>>>>>> 79b55a729d32439c8865d125c5c4c0c1a20e34a6
});
