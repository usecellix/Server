import { IsArray, IsIn, IsOptional, IsString } from 'class-validator';

/**
 * Body for `POST /excel-ai/conversation/continue` — STEPWISE_EXECUTION.md §3.
 *
 * Deliberately carries no `userId`: ownership is resolved from the session and
 * checked against the run's own `userId` (TASKS.md #170's rule), so holding a
 * runId is not by itself authority to drive that run.
 */
export class ContinueRunDto {
  @IsString()
  runId!: string;

  /**
   * The client's call on the wave that was just emitted. `rejected`/`skipped`
   * both cascade-skip dependent subtasks (SD-4) — they differ only in what the
   * closing summary reports.
   */
  @IsIn(['accepted', 'rejected', 'skipped'])
  decision!: 'accepted' | 'rejected' | 'skipped';

  /**
   * Observed post-apply state of the sheets the accepted wave touched, so the
   * next wave plans against reality rather than the shadow workbook's
   * prediction. Optional — absent readback degrades to predicted state, never
   * an error.
   */
  @IsOptional()
  @IsArray()
  readback?: unknown[];
}
