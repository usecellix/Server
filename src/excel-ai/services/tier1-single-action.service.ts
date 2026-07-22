import { Injectable, Logger } from '@nestjs/common';
import { WorkbookContext } from '../../agents/types/agent.types';
import { AppConfigService } from '../../config/app-config.service';
import { parseExecutorPayload } from '../../agents/utils/parse-agent-json.util';
import { buildTier1SystemPrompt, buildTier1UserMessage } from '../prompts/tier1-action-prompt';
import { SheetAction } from '../types/sheet-actions.types';
import { NUMERIC_FINANCIAL_HINT } from '../utils/complexity-classifier.util';
import { normalizeTier1ConditionalFormatActions } from '../utils/format-matching-rows.util';
import { extractJsonFromLlmText } from '../utils/parse-llm-response.util';
import { OpenRouterService } from './openrouter.service';

export interface Tier1ExecuteResult {
  actions: SheetAction[];
  answer: string;
  model?: string;
}

@Injectable()
export class Tier1SingleActionService {
  private readonly logger = new Logger(Tier1SingleActionService.name);

  constructor(
    private readonly openRouter: OpenRouterService,
    private readonly config: AppConfigService,
  ) {}

  async execute(
    message: string,
    actionHint: string,
    workbookContext: WorkbookContext,
  ): Promise<Tier1ExecuteResult> {
    if (actionHint === 'FIND_REPLACE' && NUMERIC_FINANCIAL_HINT.test(message)) {
      this.logger.error(
        `Classifier bug: numeric/financial FIND_REPLACE reached Tier1 — message="${message.slice(0, 120)}"`,
      );
      throw new Error('numeric_find_replace_escalation_required');
    }

    const model = this.config.openRouterModelLow;
    const raw = await this.openRouter.complete({
      systemPrompt: buildTier1SystemPrompt(actionHint),
      userMessage: buildTier1UserMessage(message, actionHint, workbookContext),
      tier: 'low',
      model,
      temperature: 0,
      maxTokens: 512,
      reasoningEffort: 'none',
    });

    const parsed = parseExecutorPayload(raw);
    const envelope = extractJsonFromLlmText(raw);
    let actions = Array.isArray(parsed?.actions) ? (parsed.actions as SheetAction[]) : [];
    const answer =
      typeof envelope?.answer === 'string' && envelope.answer.trim().length > 0
        ? envelope.answer.trim()
        : 'Applied your requested change.';

    if (actions.length !== 1) {
      this.logger.warn(
        `Tier1 expected exactly one action — got ${actions.length} for hint=${actionHint}`,
      );
    }

    actions = actions.slice(0, 1);
    if (actionHint === 'CONDITIONAL_FORMAT') {
      actions = normalizeTier1ConditionalFormatActions(actions, workbookContext, message);
    }

    return {
      actions,
      answer,
      model,
    };
  }
}
