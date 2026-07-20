import { Type } from 'class-transformer';
import {
  Allow,
  IsArray,
  IsBoolean,
  IsIn,
  IsNumber,
  IsOptional,
  IsString,
  MaxLength,
  ValidateNested,
} from 'class-validator';
import { AssistantMode } from '../types/sheet-actions.types';

export class ConversationHistoryEntryDto {
  @IsIn(['user', 'assistant'])
  role!: 'user' | 'assistant';

  @IsString()
  @MaxLength(5000)
  content!: string;
}

export class SheetCompressionDto {
  @IsNumber()
  originalRowCount!: number;

  @IsNumber()
  compressedRowCount!: number;

  @IsBoolean()
  truncated!: boolean;

  @IsBoolean()
  onDemandFetchEnabled!: boolean;

  @IsOptional()
  @IsArray()
  includedRowIndices?: number[];
}

export class SheetSnapshotDto {
  @IsString()
  sheetName!: string;

  @IsString()
  usedRange!: string;

  @IsNumber()
  rowCount!: number;

  @IsNumber()
  colCount!: number;

  @IsArray()
  @IsString({ each: true })
  headers!: string[];

  @IsOptional()
  @IsArray()
  sampleData?: unknown[][];
}

export class RichWorkbookContextDto {
  @IsString()
  activeSheet!: string;

  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => SheetSnapshotDto)
  sheets!: SheetSnapshotDto[];
}

export class ConversationContextMessageDto {
  @IsIn(['user', 'assistant'])
  role!: 'user' | 'assistant';

  @IsString()
  @MaxLength(5000)
  content!: string;

  @IsOptional()
  @IsString()
  timestamp?: string;

  @IsOptional()
  @IsIn(['question', 'answer', 'command', 'clarification'])
  type?: 'question' | 'answer' | 'command' | 'clarification';
}

export class ConversationContextDto {
  @IsOptional()
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => ConversationContextMessageDto)
  previousMessages?: ConversationContextMessageDto[];
}

export class WorkbookContextDto {
  @IsOptional()
  @IsString()
  activeSheet?: string;

  @IsOptional()
  @IsArray()
  @IsString({ each: true })
  sheets?: string[];
}

export class ConversationRequestDto {
  @IsOptional()
  @IsString()
  conversationId?: string;

  @IsString()
  @MaxLength(5000)
  message!: string;

  @IsArray()
  sheetData!: unknown[][];

  @IsOptional()
  @ValidateNested()
  @Type(() => SheetCompressionDto)
  sheetCompression?: SheetCompressionDto;

  @IsOptional()
  @ValidateNested()
  @Type(() => ConversationContextDto)
  context?: ConversationContextDto;

  /** Legacy `{ sheets: string[] }` or rich `{ sheets: SheetSnapshot[] }` from the add-in. */
  @IsOptional()
  @Allow()
  workbookContext?: WorkbookContextDto | RichWorkbookContextDto;

  @IsOptional()
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => ConversationHistoryEntryDto)
  conversationHistory?: ConversationHistoryEntryDto[];

  @IsOptional()
  @IsBoolean()
  previewEnabled?: boolean;

  /** Compressed workbook context string for LLM prompts (from add-in deep reader). */
  @IsOptional()
  @IsString()
  @MaxLength(50000)
  promptContext?: string;

  /** Quick-edit: refine a prior change set without a full workbook read. */
  @IsOptional()
  @IsString()
  refinementChangeSetId?: string;

  /** Operational mode: ask (read-only), plan (preview only), act/action (default). */
  @IsOptional()
  @IsIn(['ask', 'action', 'act', 'plan'])
  mode?: AssistantMode | 'act';
}
