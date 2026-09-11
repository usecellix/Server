import { IsString, MaxLength, MinLength } from 'class-validator';

/** Body for `PATCH /excel-ai/conversation/:conversationId` (TASKS.md #177). */
export class RenameConversationDto {
  @IsString()
  @MinLength(1)
  @MaxLength(200)
  title!: string;
}
