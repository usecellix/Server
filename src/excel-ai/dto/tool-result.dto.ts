import { IsArray, IsOptional, IsString, MaxLength } from 'class-validator';

export class ToolResultDto {
  @IsString()
  conversationId!: string;

  @IsString()
  requestId!: string;

  @IsString()
  tool!: string;

  @IsOptional()
  @IsArray()
  values?: unknown[][];

  @IsOptional()
  @IsString()
  @MaxLength(2000)
  error?: string;
}
