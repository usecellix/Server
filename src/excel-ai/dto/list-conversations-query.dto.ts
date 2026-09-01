import { Type } from 'class-transformer';
import { IsInt, IsISO8601, IsOptional, IsString, Max, MaxLength, Min } from 'class-validator';

/**
 * Query params for `GET /excel-ai/conversations` (TASKS.md #171).
 *
 * `limit` is `@Type(() => Number)`-coerced because query strings arrive as text
 * and the global pipe runs with `forbidNonWhitelisted`, which would otherwise
 * reject a perfectly valid `?limit=25`.
 */
export class ListConversationsQueryDto {
  @IsOptional()
  @Type(() => Number)
  @IsInt()
  @Min(1)
  @Max(50)
  limit?: number;

  /**
   * `updatedAt` of the last row of the previous page, ISO-8601. Cursor-based
   * rather than offset-based so a conversation updated mid-scroll can't shift
   * rows across a page boundary and cause a duplicate or a skipped entry.
   */
  @IsOptional()
  @IsISO8601()
  cursor?: string;

  /** Optional per-workbook filter — see the open #173 decision. */
  @IsOptional()
  @IsString()
  @MaxLength(200)
  workbookId?: string;
}
