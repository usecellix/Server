import { IsOptional, IsString, MaxLength, MinLength } from 'class-validator';

export class WebChatAskDto {
  @IsString()
  @MinLength(1)
  @MaxLength(4000)
  question!: string;

  /**
   * Scopes the answer to one Excel conversation the user has opened. Omitted
   * means "across my recent sessions". Ownership is enforced server-side, so
   * a conversationId belonging to another user simply yields no context
   * rather than reading their transcript.
   */
  @IsOptional()
  @IsString()
  @MaxLength(200)
  conversationId?: string;
}

export class WebChatEstimateDto {
  @IsString()
  @MinLength(1)
  @MaxLength(4000)
  question!: string;

  @IsOptional()
  @IsString()
  @MaxLength(200)
  conversationId?: string;
}
