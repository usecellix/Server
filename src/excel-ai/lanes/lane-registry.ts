import { uniquifyOutputSheets } from '../utils/output-sheet.util';
import { applyAdviceGuardrail } from '../utils/advice-guardrail.util';
import { DeterministicLane, LaneContext, LaneResult } from './lane.types';
import { STAGE2_LANES } from './stage2.lanes';
import { STAGE3_LANES } from './stage3.lanes';
import { STAGE4_LANES } from './stage4.lanes';
import { RECON_LANES } from './recon.lanes';
import { COMPLIANCE_LANES } from './compliance.lanes';

/**
 * Order matters: the first lane whose `detect` matches handles the prompt. Narrow,
 * specific lanes come before broad ones. A lane returning null falls through to
 * the next match, and then to the normal LLM route.
 */
export const DETERMINISTIC_LANES: DeterministicLane[] = [...RECON_LANES, ...STAGE4_LANES, ...COMPLIANCE_LANES, ...STAGE2_LANES, ...STAGE3_LANES];

export function matchingLanes(message: string, writeAllowed: boolean): DeterministicLane[] {
  if (!message?.trim()) return [];
  return DETERMINISTIC_LANES.filter((lane) => (writeAllowed || !lane.writes) && lane.detect(message));
}

export async function runDeterministicLanes(
  ctx: LaneContext,
): Promise<{ lane: DeterministicLane; result: LaneResult } | null> {
  for (const lane of matchingLanes(ctx.message, ctx.writeAllowed)) {
    const result = await lane.run(ctx);
    if (!result) continue;
    if (result.kind === 'actions') {
      if (!ctx.writeAllowed) continue; // never propose a write outside Action mode
      return {
        lane,
        result: {
          ...result,
          actions: uniquifyOutputSheets(result.actions, ctx.sheetNames),
          answer: applyAdviceGuardrail(result.answer, ctx.message).answer,
        },
      };
    }
    return { lane, result: { ...result, answer: applyAdviceGuardrail(result.answer, ctx.message).answer } };
  }
  return null;
}
