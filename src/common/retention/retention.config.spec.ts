import { RETENTION, expiryFromNow } from './retention.config';
import { ChangeSetSchema } from '../../audit/schemas/change-set.schema';
import { AiPromptSchema } from '../../llm-usage/schemas/ai-prompt.schema';

describe('P0.3 retention limits', () => {
  const env = { ...process.env };
  afterEach(() => {
    process.env = { ...env };
  });

  it('defaults: change sets 7 days, prompts 24 h, planner disk log off', () => {
    delete process.env.CHANGESET_TTL_HOURS;
    delete process.env.AI_PROMPT_TTL_HOURS;
    delete process.env.PLANNER_FILE_LOG;
    expect(RETENTION.changeSetHours()).toBe(168);
    expect(RETENTION.aiPromptHours()).toBe(24);
    expect(RETENTION.plannerFileLogEnabled()).toBe(false);
  });

  it('is configurable and ignores junk values', () => {
    process.env.CHANGESET_TTL_HOURS = '2';
    process.env.AI_PROMPT_TTL_HOURS = 'abc';
    process.env.PLANNER_FILE_LOG = 'on';
    expect(RETENTION.changeSetHours()).toBe(2);
    expect(RETENTION.aiPromptHours()).toBe(24);
    expect(RETENTION.plannerFileLogEnabled()).toBe(true);
  });

  it('expiry is in the future by the given hours', () => {
    const t = expiryFromNow(1).getTime() - Date.now();
    expect(t).toBeGreaterThan(3_590_000);
    expect(t).toBeLessThan(3_610_000);
  });

  it.each([
    ['change_sets', ChangeSetSchema],
    ['ai_prompts', AiPromptSchema],
  ])('%s has a TTL index on expiresAt', (_name, schema) => {
    const ttl = schema.indexes().find(([keys, opts]) => 'expiresAt' in keys && opts?.expireAfterSeconds === 0);
    expect(ttl).toBeDefined();
  });
});

describe('P0.3 request log keeps no values by default', () => {
  const env = { ...process.env };
  afterEach(() => {
    process.env = { ...env };
  });
  it('bodies are off unless REQUEST_LOG_BODIES=on', () => {
    delete process.env.REQUEST_LOG_BODIES;
    expect(RETENTION.requestLogBodiesEnabled()).toBe(false);
    process.env.REQUEST_LOG_BODIES = 'on';
    expect(RETENTION.requestLogBodiesEnabled()).toBe(true);
  });
});
