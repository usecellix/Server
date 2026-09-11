import { ConfigService } from '@nestjs/config';
import { AppConfigService } from '../src/config/app-config.service';

function buildConfig(env: Record<string, string>): AppConfigService {
  return new AppConfigService(new ConfigService(env));
}

describe('AppConfigService — MODEL_PROFILE', () => {
  it('defaults to prod when MODEL_PROFILE is unset, leaving every tier model exactly as configured', () => {
    const config = buildConfig({
      OPENROUTER_MODEL_LOW: 'openai/gpt-5-nano',
      OPENROUTER_MODEL_MEDIUM: 'openai/gpt-5-mini',
      OPENROUTER_MODEL_HIGH: 'openai/gpt-5',
    });

    expect(config.modelProfile).toBe('prod');
    expect(config.openRouterModelLow).toBe('openai/gpt-5-nano');
    expect(config.openRouterModelMedium).toBe('openai/gpt-5-mini');
    expect(config.openRouterModelHigh).toBe('openai/gpt-5');
    expect(config.openRouterModelTier1).toBe('openai/gpt-5-mini');
  });

  it('MODEL_PROFILE=dev swaps every tier to CELLIX_DEV_MODEL regardless of the prod env vars', () => {
    const config = buildConfig({
      MODEL_PROFILE: 'dev',
      OPENROUTER_MODEL_LOW: 'openai/gpt-5-nano',
      OPENROUTER_MODEL_MEDIUM: 'openai/gpt-5-mini',
      OPENROUTER_MODEL_HIGH: 'openai/gpt-5',
      OPENROUTER_MODEL_TIER1: 'openai/gpt-5-mini',
    });

    expect(config.modelProfile).toBe('dev');
    expect(config.openRouterModelLow).toBe('z-ai/glm-5.3-flash');
    expect(config.openRouterModelMedium).toBe('z-ai/glm-5.3-flash');
    expect(config.openRouterModelHigh).toBe('z-ai/glm-5.3-flash');
    expect(config.openRouterModelTier1).toBe('z-ai/glm-5.3-flash');
  });

  it('CELLIX_DEV_MODEL is overridable and defaults to z-ai/glm-5.3-flash', () => {
    expect(buildConfig({}).cellixDevModel).toBe('z-ai/glm-5.3-flash');
    expect(
      buildConfig({ CELLIX_DEV_MODEL: 'some/other-model' }).cellixDevModel,
    ).toBe('some/other-model');
  });

  it('a custom CELLIX_DEV_MODEL is what dev profile swaps every tier to', () => {
    const config = buildConfig({
      MODEL_PROFILE: 'dev',
      CELLIX_DEV_MODEL: 'some/other-model',
    });

    expect(config.openRouterModelLow).toBe('some/other-model');
    expect(config.openRouterModelHigh).toBe('some/other-model');
    expect(config.openRouterModelTier1).toBe('some/other-model');
  });

  it('an unrecognized MODEL_PROFILE value falls back to prod rather than silently matching dev', () => {
    const config = buildConfig({ MODEL_PROFILE: 'staging' });
    expect(config.modelProfile).toBe('prod');
  });
});
