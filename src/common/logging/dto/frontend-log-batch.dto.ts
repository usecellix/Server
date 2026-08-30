import { Type } from 'class-transformer';
import {
  ArrayMaxSize,
  ArrayMinSize,
  IsArray,
  IsIn,
  IsObject,
  IsOptional,
  IsString,
  MaxLength,
  ValidateNested,
} from 'class-validator';

const LEVELS = ['error', 'warn', 'info', 'action'] as const;
/**
 * MUST stay in sync with `FrontendTelemetryCategory` in
 * `client/src/services/frontendTelemetry.ts`. These are two independently
 * maintained copies of one contract — the AD-7 drift pattern again — and the
 * failure mode is nastier than it looks: a single unknown category 400s the
 * WHOLE batch, so unrelated events posted alongside it are silently lost too.
 *
 * That is exactly what happened when `'verify'` was added client-side for
 * TASKS.md #150 and not here: the apply itself succeeded (`/audit/apply` 201),
 * but `accept.success` and every verification result vanished with the rejected
 * batch, leaving the run looking like it had hung. TASKS.md #159.
 */
const CATEGORIES = [
  'console',
  'preview',
  'accept',
  'reject',
  'apply',
  'verify',
  'sse',
  'navigation',
  'other',
] as const;

/** Exported so the parity test can compare against the client's own union. */
export const TELEMETRY_CATEGORIES = CATEGORIES;

export class FrontendLogEventDto {
  @IsOptional()
  @IsString()
  @MaxLength(40)
  ts?: string;

  @IsIn(LEVELS)
  level!: (typeof LEVELS)[number];

  @IsIn(CATEGORIES)
  category!: (typeof CATEGORIES)[number];

  @IsString()
  @MaxLength(120)
  event!: string;

  @IsString()
  @MaxLength(2000)
  message!: string;

  @IsOptional()
  @IsString()
  @MaxLength(120)
  conversationId?: string;

  @IsOptional()
  @IsString()
  @MaxLength(120)
  changeSetId?: string;

  @IsOptional()
  @IsString()
  @MaxLength(120)
  sessionId?: string;

  @IsOptional()
  @IsString()
  @MaxLength(200)
  workbookKey?: string;

  @IsOptional()
  @IsString()
  @MaxLength(500)
  userAgent?: string;

  @IsOptional()
  @IsString()
  @MaxLength(500)
  pageUrl?: string;

  @IsOptional()
  @IsObject()
  details?: Record<string, unknown>;
}

export class FrontendLogBatchDto {
  @IsArray()
  @ArrayMinSize(1)
  @ArrayMaxSize(50)
  @ValidateNested({ each: true })
  @Type(() => FrontendLogEventDto)
  events!: FrontendLogEventDto[];
}
