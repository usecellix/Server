import { Type } from 'class-transformer';
import { IsInt, IsISO8601, IsOptional, Max, Min } from 'class-validator';

/** Query params for `GET /billing/ledger` (CREDIT_SYSTEM_SCHEMA.md §5). */
export class ListLedgerQueryDto {
  @IsOptional()
  @Type(() => Number)
  @IsInt()
  @Min(1)
  @Max(100)
  limit?: number;

  /** `createdAt` of the last row of the previous page, ISO-8601. */
  @IsOptional()
  @IsISO8601()
  cursor?: string;
}
