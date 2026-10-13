import { Controller, Get, NotFoundException, Param, Query, UseGuards } from '@nestjs/common';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { AdminGuard } from './admin.guard';
import { AdminUsersService, UserSort } from './admin-users.service';
import { AdminPromptsService, PromptSort } from './admin-prompts.service';
import { AdminOverviewService } from './admin-overview.service';
import { AdminBillingService, LedgerSort, SubSort } from './admin-billing.service';
import { AdminModelsService, UsageSort } from './admin-models.service';
import { resolveAdminRange } from './admin-range.util';

const USER_SORTS: UserSort[] = ['recent', 'spend', 'prompts', 'credits', 'creditsUsed', 'seen'];
const PROMPT_SORTS: PromptSort[] = ['recent', 'cost', 'tokens', 'calls'];
const SUB_SORTS: SubSort[] = ['recent', 'price', 'period'];
const LEDGER_SORTS: LedgerSort[] = ['recent', 'amount'];
const USAGE_SORTS: UsageSort[] = ['cost', 'calls', 'errors', 'latency', 'name'];
const PROMPT_STATUSES = ['error', 'ok', 'running'] as const;

function page(value: string | undefined): number {
  return Math.max(1, Number(value) || 1);
}

function pick<T extends string>(value: string | undefined, allowed: readonly T[], fallback: T): T {
  return (allowed as readonly string[]).includes(value ?? '') ? (value as T) : fallback;
}

/**
 * Read surface for the Dashboard admin app (AdminGuard's docblock explains
 * the auth model). Each route mirrors one `Dashboard/src/lib/data/*.ts`
 * function's filters/shape 1:1 so the Dashboard's own functions become thin
 * fetch wrappers with unchanged signatures — see TASKS.md admin-api-migration.
 */
@UseGuards(AdminGuard)
@Controller('admin')
export class AdminController {
  constructor(
    private readonly adminUsers: AdminUsersService,
    private readonly adminPrompts: AdminPromptsService,
    private readonly adminOverview: AdminOverviewService,
    private readonly adminBilling: AdminBillingService,
    private readonly adminModels: AdminModelsService,
  ) {}

  @Get('users')
  @SkipEnvelope()
  async listUsers(
    @Query('q') q?: string,
    @Query('page') pageParam?: string,
    @Query('sort') sort?: string,
    @Query('plan') plan?: string,
    @Query('range') range?: string,
  ) {
    return this.adminUsers.listUsers({
      q,
      plan,
      sort: pick(sort, USER_SORTS, 'recent'),
      page: page(pageParam),
      since: resolveAdminRange(range).from ?? undefined,
    });
  }

  @Get('users/:id')
  @SkipEnvelope()
  async getUser(@Param('id') id: string) {
    const user = await this.adminUsers.getUser(id);
    if (!user) throw new NotFoundException('USER_NOT_FOUND');
    return user;
  }

  @Get('prompts')
  @SkipEnvelope()
  async listPrompts(
    @Query('range') range?: string,
    @Query('q') q?: string,
    @Query('userId') userId?: string,
    @Query('status') status?: string,
    @Query('sort') sort?: string,
    @Query('page') pageParam?: string,
  ) {
    return this.adminPrompts.listPrompts({
      since: resolveAdminRange(range).from ?? undefined,
      userId,
      status: (PROMPT_STATUSES as readonly string[]).includes(status ?? '') ? (status as (typeof PROMPT_STATUSES)[number]) : undefined,
      q,
      sort: pick(sort, PROMPT_SORTS, 'recent'),
      page: page(pageParam),
    });
  }

  @Get('prompts/:promptId')
  @SkipEnvelope()
  async getPrompt(@Param('promptId') promptId: string) {
    const detail = await this.adminPrompts.getPrompt(promptId);
    if (!detail) throw new NotFoundException('PROMPT_NOT_FOUND');
    return detail;
  }

  @Get('overview')
  @SkipEnvelope()
  async getOverview(@Query('range') range?: string) {
    const resolved = resolveAdminRange(range);
    const overview = await this.adminOverview.getOverview({ from: resolved.from, to: resolved.to, bucket: resolved.bucket });
    if (!overview) throw new NotFoundException('OVERVIEW_UNAVAILABLE');
    return overview;
  }

  @Get('billing/summary')
  @SkipEnvelope()
  async getBillingSummary(@Query('range') range?: string) {
    const summary = await this.adminBilling.getBillingSummary(resolveAdminRange(range).from);
    if (!summary) throw new NotFoundException('BILLING_UNAVAILABLE');
    return summary;
  }

  @Get('billing/subscriptions')
  @SkipEnvelope()
  async listSubscriptions(
    @Query('range') range?: string,
    @Query('q') q?: string,
    @Query('status') status?: string,
    @Query('plan') plan?: string,
    @Query('sort') sort?: string,
    @Query('page') pageParam?: string,
  ) {
    return this.adminBilling.listSubscriptions({
      since: resolveAdminRange(range).from,
      q,
      status,
      plan,
      sort: pick(sort, SUB_SORTS, 'recent'),
      page: page(pageParam),
    });
  }

  @Get('billing/ledger')
  @SkipEnvelope()
  async listLedger(
    @Query('range') range?: string,
    @Query('q') q?: string,
    @Query('entryType') entryType?: string,
    @Query('sort') sort?: string,
    @Query('page') pageParam?: string,
  ) {
    return this.adminBilling.listLedger({
      since: resolveAdminRange(range).from,
      q,
      entryType,
      sort: pick(sort, LEDGER_SORTS, 'recent'),
      page: page(pageParam),
    });
  }

  @Get('models')
  @SkipEnvelope()
  async getModelUsage(
    @Query('range') range?: string,
    @Query('q') q?: string,
    @Query('sort') sort?: string,
    @Query('modelPage') modelPage?: string,
    @Query('callerPage') callerPage?: string,
  ) {
    const usage = await this.adminModels.getModelUsage(resolveAdminRange(range).from, {
      q,
      sort: pick(sort, USAGE_SORTS, 'cost'),
      modelPage: page(modelPage),
      callerPage: page(callerPage),
    });
    if (!usage) throw new NotFoundException('MODELS_UNAVAILABLE');
    return usage;
  }
}
