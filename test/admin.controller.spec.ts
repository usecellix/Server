import { AdminController } from '../src/admin/admin.controller';

function buildController() {
  const adminUsers = { listUsers: jest.fn().mockResolvedValue({ rows: [], total: 0 }), getUser: jest.fn() };
  const adminPrompts = { listPrompts: jest.fn().mockResolvedValue({ rows: [], total: 0 }), getPrompt: jest.fn() };
  const adminOverview = { getOverview: jest.fn().mockResolvedValue({ totals: {} }) };
  const adminBilling = {
    getBillingSummary: jest.fn().mockResolvedValue({ activeCount: 0 }),
    listSubscriptions: jest.fn().mockResolvedValue({ rows: [], total: 0 }),
    listLedger: jest.fn().mockResolvedValue({ rows: [], total: 0 }),
  };
  const adminModels = { getModelUsage: jest.fn().mockResolvedValue({ byModel: [] }) };
  const controller = new AdminController(
    adminUsers as never,
    adminPrompts as never,
    adminOverview as never,
    adminBilling as never,
    adminModels as never,
  );
  return { controller, adminUsers, adminPrompts, adminOverview, adminBilling, adminModels };
}

describe('AdminController', () => {
  it('listUsers falls back to a safe sort for an unrecognized value rather than passing it through', async () => {
    const { controller, adminUsers } = buildController();
    await controller.listUsers(undefined, undefined, 'not-a-real-sort', undefined, undefined);
    expect(adminUsers.listUsers).toHaveBeenCalledWith(expect.objectContaining({ sort: 'recent' }));
  });

  it('listUsers resolves range=7d into a since date about 7 days back', async () => {
    const { controller, adminUsers } = buildController();
    await controller.listUsers(undefined, undefined, undefined, undefined, '7d');
    const since = adminUsers.listUsers.mock.calls[0][0].since as Date;
    expect(Date.now() - since.getTime()).toBeGreaterThan(6.9 * 86_400_000);
    expect(Date.now() - since.getTime()).toBeLessThan(7.1 * 86_400_000);
  });

  it('listUsers with range=all passes no since filter', async () => {
    const { controller, adminUsers } = buildController();
    await controller.listUsers(undefined, undefined, undefined, undefined, 'all');
    expect(adminUsers.listUsers.mock.calls[0][0].since).toBeUndefined();
  });

  it('getUser 404s when the service returns null', async () => {
    const { controller, adminUsers } = buildController();
    adminUsers.getUser.mockResolvedValue(null);
    await expect(controller.getUser('missing')).rejects.toThrow('USER_NOT_FOUND');
  });

  it('getPrompt 404s when the service returns null', async () => {
    const { controller, adminPrompts } = buildController();
    adminPrompts.getPrompt.mockResolvedValue(null);
    await expect(controller.getPrompt('missing')).rejects.toThrow('PROMPT_NOT_FOUND');
  });

  it('listPrompts drops an unrecognized status instead of passing it to the query', async () => {
    const { controller, adminPrompts } = buildController();
    await controller.listPrompts(undefined, undefined, undefined, 'not-a-real-status', undefined, undefined);
    expect(adminPrompts.listPrompts).toHaveBeenCalledWith(expect.objectContaining({ status: undefined }));
  });

  it('page numbers below 1 (or non-numeric) clamp to 1', async () => {
    const { controller, adminUsers } = buildController();
    await controller.listUsers(undefined, '-5', undefined, undefined, undefined);
    expect(adminUsers.listUsers).toHaveBeenCalledWith(expect.objectContaining({ page: 1 }));
  });

  it('getOverview 404s when the service has no db (returns null)', async () => {
    const { controller, adminOverview } = buildController();
    adminOverview.getOverview.mockResolvedValue(null);
    await expect(controller.getOverview(undefined)).rejects.toThrow('OVERVIEW_UNAVAILABLE');
  });

  it('getModelUsage passes through q/sort/pagination', async () => {
    const { controller, adminModels } = buildController();
    await controller.getModelUsage('90d', 'gpt', 'errors', '2', '3');
    expect(adminModels.getModelUsage).toHaveBeenCalledWith(
      expect.any(Date),
      expect.objectContaining({ q: 'gpt', sort: 'errors', modelPage: 2, callerPage: 3 }),
    );
  });
});
