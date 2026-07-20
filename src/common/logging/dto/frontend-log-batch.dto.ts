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
const CATEGORIES = [
  'console',
  'preview',
  'accept',
  'reject',
  'apply',
  'sse',
  'navigation',
  'other',
] as const;

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
