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

  /**
   * "prod" (default) leaves every tier's model assignment exactly as
   * configured. "dev" makes every tier resolve to `cellixDevModel` instead —
   * a single override switch for local/dev runs, without touching any
   * tier-routing logic (ModelRouter, OpenRouterService, and each agent all
   * read their model through this service, so the swap happens once here).
   */
  get modelProfile(): 'prod' | 'dev' {
    return this.configService.get<string>('MODEL_PROFILE', 'prod') === 'dev' ? 'dev' : 'prod';
  }

  get cellixDevModel(): string {
    return this.configService.get<string>('CELLIX_DEV_MODEL', 'z-ai/glm-5.3-flash');
  }

  /** @deprecated Use tier models (LOW / MEDIUM / HIGH). Kept for OpenRouterService default only. */
  get openRouterModel(): string {
    return this.openRouterModelMedium;
  }

  get openRouterModelLow(): string {
    if (this.modelProfile === 'dev') return this.cellixDevModel;
    return this.configService.get<string>('OPENROUTER_MODEL_LOW', 'openai/gpt-5-mini');
  }

  get openRouterModelMedium(): string {
    if (this.modelProfile === 'dev') return this.cellixDevModel;
    return this.configService.get<string>('OPENROUTER_MODEL_MEDIUM', 'openai/gpt-5-mini');
  }

  get openRouterModelHigh(): string {
    if (this.modelProfile === 'dev') return this.cellixDevModel;
    return this.configService.get<string>('OPENROUTER_MODEL_HIGH', 'openai/gpt-5');
  }

  /**
   * Tier 1 emits real SheetAction JSON that writes to cells directly, skipping
   * the Planner/Verifier pipeline that would catch a malformed or mis-ranged
   * action. Deliberately decoupled from the shared LOW tier (TASKS.md #162) —
   * LOW is safe on the cheapest model because the router only picks a route
   * label, but a bad Tier 1 write is the #87 failure class. Defaults to the
   * MEDIUM model; override independently via OPENROUTER_MODEL_TIER1 if needed.
   */
  get openRouterModelTier1(): string {
    if (this.modelProfile === 'dev') return this.cellixDevModel;
    return this.configService.get<string>('OPENROUTER_MODEL_TIER1', this.openRouterModelMedium);
  }

  /**
   * Spec 16 fix #2 — model-selection experiment for Planner calls specifically.
   * `openRouterModelHigh` (gpt-5, a reasoning model) is the Planner's job: pure
   * JSON decomposition, not open-ended reasoning, and reasoning tokens sharing
   * the completion budget is exactly what caused the Spec 16 empty-response
   * incident (completionTokens === reasoningTokens, zero content emitted).
   * Decoupled from HIGH so a non-reasoning/lighter-reasoning model (e.g. a
   * GPT-4.1/GPT-4o-class model with strong structured-output adherence) can be
   * evaluated for the Planner in isolation — without changing the Executor's
   * model, which does its own separate reasoning over tool calls. Defaults to
   * HIGH, so this is a no-op until OPENROUTER_MODEL_PLANNER is explicitly set.
   */
  get openRouterModelPlanner(): string {
    if (this.modelProfile === 'dev') return this.cellixDevModel;
    return this.configService.get<string>('OPENROUTER_MODEL_PLANNER', this.openRouterModelHigh);
  }

  /**
   * LlmRouterService's LOW-tier classification call. Decoupled from the
   * shared LOW tier — `multi-sheet.service.ts`'s summary call and
   * conversation.service.ts's ambiguity-clarification quickCall() also read
   * openRouterModelLow, and a router-specific model eval must not silently
   * change those too. Defaults to openRouterModelLow, so this is a no-op
   * until OPENROUTER_MODEL_ROUTER is explicitly set.
   */
  get openRouterModelRouter(): string {
    if (this.modelProfile === 'dev') return this.cellixDevModel;
    return this.configService.get<string>('OPENROUTER_MODEL_ROUTER', this.openRouterModelLow);
  }

  /**
   * Tier2GenerateVerifyService's generate-pass call into ExecutorAgent.
   * ExecutorAgent.execute() is shared verbatim with Tier 3's Executor
   * (same class, same default model) — this override lets Tier 2's call
   * pass a different model without touching Tier 3, which must stay on
   * openRouterModelHigh unconditionally. Defaults to openRouterModelHigh,
   * so this is a no-op until OPENROUTER_MODEL_TIER2_GENERATE is explicitly
   * set.
   */
  get openRouterModelTier2Generate(): string {
    if (this.modelProfile === 'dev') return this.cellixDevModel;
    return this.configService.get<string>(
      'OPENROUTER_MODEL_TIER2_GENERATE',
      this.openRouterModelHigh,
    );
  }

  /**
   * Lets eval/run-live-eval.ts authenticate without a browser session.
   * AuthGuard only honors this outside production — see auth.guard.ts.
   */
  get evalBypassToken(): string | undefined {
    const value = this.configService.get<string>('CELLIX_EVAL_BYPASS_TOKEN', '');
    return value?.trim() ? value.trim() : undefined;
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

  get razorpayKeyId(): string | undefined {
    const value = this.configService.get<string>('RAZORPAY_KEY_ID', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get razorpayKeySecret(): string | undefined {
    const value = this.configService.get<string>('RAZORPAY_KEY_SECRET', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get razorpayWebhookSecret(): string | undefined {
    const value = this.configService.get<string>('RAZORPAY_WEBHOOK_SECRET', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get razorpayPlanIdSolo(): string | undefined {
    const value = this.configService.get<string>('RAZORPAY_PLAN_ID_SOLO', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get razorpayPlanIdFirm(): string | undefined {
    const value = this.configService.get<string>('RAZORPAY_PLAN_ID_FIRM', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get razorpayPlanIdBeta(): string | undefined {
    const value = this.configService.get<string>('RAZORPAY_PLAN_ID_BETA', '');
    return value?.trim() ? value.trim() : undefined;
  }

  get checkoutSuccessUrl(): string {
    return this.configService.get<string>('CHECKOUT_SUCCESS_URL', 'http://localhost:5173/checkout/success');
  }

  get checkoutCancelUrl(): string {
    return this.configService.get<string>('CHECKOUT_CANCEL_URL', 'http://localhost:5173/checkout');
  }
}
