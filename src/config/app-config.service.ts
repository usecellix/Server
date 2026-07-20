import { Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';

@Injectable()
export class AppConfigService {
  constructor(private readonly configService: ConfigService) {}

  get nodeEnv(): string {
    return this.configService.get<string>('NODE_ENV', 'development');
  }

  get port(): number {
    return this.configService.get<number>('PORT', 4001);
  }

  get mongoUrl(): string {
    return this.configService.get<string>('MONGODB_URL', 'mongodb://127.0.0.1:27017/cellix');
  }

  get mongoDbName(): string {
    return this.configService.get<string>('MONGODB_DB_NAME', 'cellix');
  }

  get openRouterApiKey(): string | undefined {
    const key = this.configService.get<string>('OPENROUTER_API_KEY', '');
    return key?.trim() ? key.trim() : undefined;
  }

  /** @deprecated Use tier models (LOW / MEDIUM / HIGH). Kept for OpenRouterService default only. */
  get openRouterModel(): string {
    return this.openRouterModelMedium;
  }

  get openRouterModelLow(): string {
    return this.configService.get<string>('OPENROUTER_MODEL_LOW', 'openai/gpt-5-mini');
  }

  get openRouterModelMedium(): string {
    return this.configService.get<string>('OPENROUTER_MODEL_MEDIUM', 'openai/gpt-5-mini');
  }

  get openRouterModelHigh(): string {
    return this.configService.get<string>('OPENROUTER_MODEL_HIGH', 'openai/gpt-5');
  }

  get openRouterHttpReferer(): string {
    return this.configService.get<string>('OPENROUTER_HTTP_REFERER', 'https://cellix.local');
  }

  get openAiApiKey(): string | undefined {
    const key = this.configService.get<string>('OPENAI_API_KEY', '');
    return key?.trim() ? key.trim() : undefined;
  }

  get openAiModel(): string {
    return this.configService.get<string>('OPENAI_MODEL', 'gpt-4o-mini');
  }

  get openAiModelLow(): string {
    return this.configService.get<string>('OPENAI_MODEL_LOW', this.openAiModel);
  }

  get openAiModelMedium(): string {
    return this.configService.get<string>('OPENAI_MODEL_MEDIUM', this.openAiModel);
  }

  get openAiModelHigh(): string {
    return this.configService.get<string>('OPENAI_MODEL_HIGH', this.openAiModel);
  }

  /** OpenRouter is preferred when configured. */
  get usesOpenRouter(): boolean {
    return Boolean(this.openRouterApiKey);
  }

  get hasLlmProvider(): boolean {
    return Boolean(this.openRouterApiKey || this.openAiApiKey);
  }

  get betterAuthSecret(): string | undefined {
    return this.configService.get<string>('BETTER_AUTH_SECRET');
  }

  get betterAuthUrl(): string {
    return this.configService.get<string>('BETTER_AUTH_URL', this.clientOrigin);
  }

  get clientOrigin(): string {
    return this.configService.get<string>('CLIENT_ORIGIN', 'https://localhost:3000');
  }

  get googleClientId(): string | undefined {
    const value = this.configService.get<string>('GOOGLE_CLIENT_ID', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get googleClientSecret(): string | undefined {
    const value = this.configService.get<string>('GOOGLE_CLIENT_SECRET', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get microsoftClientId(): string | undefined {
    const value = this.configService.get<string>('MICROSOFT_CLIENT_ID', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get microsoftClientSecret(): string | undefined {
    const value = this.configService.get<string>('MICROSOFT_CLIENT_SECRET', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get microsoftTenantId(): string {
    return this.configService.get<string>('MICROSOFT_TENANT_ID', 'common');
  }
}
