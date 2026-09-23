import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';
import { LlmCompletionOutcome, OpenRouterService } from '../excel-ai/services/openrouter.service';
import { PlannerOutput } from './types/agent.types';
import { parseAgentJson } from './utils/parse-agent-json.util';
import {
  applyBuildSpecToSubtasks,
  BuildSpec,
  groundBuildSpec,
  parseBuildSpec,
  shouldExtractBuildSpec,
} from './utils/build-spec.util';
import { splitSpecPinnedSubtasks } from './utils/header-table-split.util';
import { addUsage, UsageTotals } from './utils/usage-accumulator.util';

const SPEC_SYSTEM_PROMPT = `You extract the column lists a user explicitly wrote in a spreadsheet request.

Return ONLY one JSON object: {"sheets":[{"names":["<sheet name>", ...],"columns":["<column>", ...]}]}

Rules:
- Include a sheet ONLY when the user spelled out its columns (e.g. "each month sheet includes Unit No, Guest, Check in ...").
- "columns" are copied VERBATIM from the request, in the order the user wrote them. Never rename, translate, merge, split, reorder, or add columns. Do not add columns the user did not list, even obvious computed ones.
- When one column list applies to several sheets (every month), put all their names in "names" (January..December, or whatever the user's request implies).
- If the user gave no explicit column list, return {"sheets":[]}.`;

/**
 * Phase 1 of LONG_PROMPT_RELIABILITY_PLAN.md — reads the user's own column
 * lists once, so the Executor is told them verbatim rather than trusting them to
 * survive prompt -> plan prose -> re-typing.
 *
 * Purely additive: every failure path returns the plan untouched. A build must
 * never be blocked, or slowed by a retry, because this optional step hiccuped.
 */
@Injectable()
export class SpecExtractorAgent {
  private readonly logger = new Logger(SpecExtractorAgent.name);

  constructor(
    private readonly llm: OpenRouterService,
    private readonly config: AppConfigService,
  ) {}

  async extract(prompt: string, usageTotals?: UsageTotals): Promise<BuildSpec | null> {
    try {
      const outcome: LlmCompletionOutcome = {};
      const raw = await this.llm.complete({
        systemPrompt: SPEC_SYSTEM_PROMPT,
        userMessage: prompt,
        model: this.config.openRouterModelHigh,
        temperature: 0,
        maxTokens: 1500,
        reasoningEffort: 'low',
        responseFormat: 'json_object',
        outcome,
      });
      addUsage(usageTotals, outcome.usage);
      if (outcome.truncated === true) return null;

      const parsed = parseBuildSpec(parseAgentJson(raw));
      return parsed ? groundBuildSpec(parsed, prompt) : null;
    } catch (error) {
      this.logger.warn(
        `Spec extraction skipped: ${error instanceof Error ? error.message : String(error)}`,
      );
      return null;
    }
  }

  /** Returns `plan` with `expectedHeaders` stamped on the subtasks that write spec'd headers. */
  async attach(
    prompt: string,
    plan: PlannerOutput,
    usageTotals?: UsageTotals,
  ): Promise<PlannerOutput> {
    if (!shouldExtractBuildSpec(prompt, plan.subtasks)) return plan;

    const spec = await this.extract(prompt, usageTotals);
    if (!spec) return plan;

    const pinned = applyBuildSpecToSubtasks(plan.subtasks, spec);
    const stamped = pinned.filter((subtask) => subtask.expectedHeaders?.length).length;
    // LONG_PROMPT_RELIABILITY_PLAN.md — split the create+headers+table piece
    // (now code-certain from the spec) off of each pinned subtask so the
    // Executor's own generation is lighter and more likely to finish within
    // its iteration budget. Two live runs showed a single month subtask
    // hitting "max iterations" or timing out running completely ALONE — not a
    // concurrency problem, a per-subtask workload problem.
    const subtasks = splitSpecPinnedSubtasks(pinned);
    this.logger.log(
      `Build spec: ${spec.sheets.length} sheet group(s), ${stamped}/${plan.subtasks.length} subtask(s) pinned to the user's columns, ${subtasks.length - plan.subtasks.length} deterministic header/table step(s) split out`,
    );
    return { ...plan, subtasks };
  }
}
