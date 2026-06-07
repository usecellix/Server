import { Type } from 'class-transformer';
import {
  IsArray,
  IsIn,
  IsOptional,
  IsString,
  MaxLength,
  ValidateNested,
} from 'class-validator';

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
  @Type(() => ConversationContextDto)
  context?: ConversationContextDto;

  @IsOptional()
  @ValidateNested()
  @Type(() => WorkbookContextDto)
  workbookContext?: WorkbookContextDto;
}
