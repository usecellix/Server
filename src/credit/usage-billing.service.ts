import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';
import type { LlmUsageContext } from '../llm-usage/llm-usage.context';
import { CreditGateService } from './credit-gate.service';
import { CreditLedgerService } from './credit-ledger.service';
import { AI_USAGE_ACTION_TYPE, type UsageDebitResult } from './types/credit.types';

export interface UsageSettlement extends UsageDebitResult {
  costUsd: number;
}

/**
 * Charges each HTTP request for the AI it actually consumed: the summed real
 * provider cost of its LLM calls (LlmUsageContext.costUsd, fed by
 * LlmUsageService.recordCall) times CREDITS_PER_USD. A request that made no
 * model call costs nothing. A stepwise build is charged wave by wave, since
 * every /continue is its own request, so the balance falls as the build runs
 * and the next wave is refused once it reaches 0.
 */
@Injectable()
export class UsageBillingService {
  private readonly logger = new Logger(UsageBillingService.name);

  constructor(
    private readonly config: AppConfigService,
    private readonly gate: CreditGateService,
    private readonly ledger: CreditLedgerService,
  ) {}

  creditsForCost(costUsd: number): number {
    if (!(costUsd > 0)) return 0;
    // Rounded to 6 dp first so float noise (0.1 + 0.2) can't tip ceil up a credit.
    return Math.ceil(Math.round(costUsd * this.config.creditsPerUsd * 1e6) / 1e6);
  }

  /** Pre-flight: a request may start only while some balance remains. */
  async canStart(userId: string): Promise<boolean> {
    return this.gate.hasAnyBalance(userId);
  }

  /**
   * Debits the not-yet-billed part of this request's cost. Safe to call more
   * than once per request (at stream end, then again at request end to catch
   * calls that finished after the stream closed): calls are chained on the
   * context, and each only bills cost accrued since the previous one.
   */
  settle(context: LlmUsageContext): Promise<UsageSettlement | null> {
    const run = (context.billing ?? Promise.resolve()).then(() => this.settleNow(context));
    context.billing = run.catch(() => undefined);
    return run;
  }

  private async settleNow(context: LlmUsageContext): Promise<UsageSettlement | null> {
    const userId = context.userId;
    const unbilled = (context.costUsd ?? 0) - (context.billedCostUsd ?? 0);
    if (!userId || !(unbilled > 0)) return null;

    const credits = this.creditsForCost(unbilled);
    context.billedCostUsd = context.costUsd;
    try {
      const result = await this.ledger.debitUsage(userId, credits, {
        conversationId: context.conversationId,
        promptId: context.promptId || undefined,
        costUsd: unbilled,
      });
      if (result.debited < result.requested) {
        this.logger.warn(
          `${AI_USAGE_ACTION_TYPE} short by ${result.requested - result.debited} credits user=${userId} prompt=${context.promptId} — balance exhausted`,
        );
      }
      return { ...result, costUsd: unbilled };
    } catch (err: unknown) {
      this.logger.error(
        `${AI_USAGE_ACTION_TYPE} debit failed user=${userId} prompt=${context.promptId} credits=${credits}: ${
          err instanceof Error ? err.message : String(err)
        }`,
      );
      return null;
    }
  }
}
