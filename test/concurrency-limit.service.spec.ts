import { ConcurrencyLimitService } from '../src/common/guards/concurrency-limit.service';

function serviceWithCap(maxConcurrentRequestsPerUser: number): ConcurrencyLimitService {
  return new ConcurrencyLimitService({ maxConcurrentRequestsPerUser } as never);
}

describe('ConcurrencyLimitService (TASKS.md #344)', () => {
  it('admits requests up to the configured cap, then refuses the next one', () => {
    const service = serviceWithCap(2);

    expect(service.tryAcquire('u1')).toBe(true);
    expect(service.tryAcquire('u1')).toBe(true);
    expect(service.tryAcquire('u1')).toBe(false); // 3rd concurrent request for the same user
  });

  it('tracks each user independently — one user being at the cap does not affect another', () => {
    const service = serviceWithCap(1);

    expect(service.tryAcquire('u1')).toBe(true);
    expect(service.tryAcquire('u2')).toBe(true); // different user, own slot
    expect(service.tryAcquire('u1')).toBe(false);
  });

  it('frees a slot on release, letting a subsequent request through', () => {
    const service = serviceWithCap(1);

    expect(service.tryAcquire('u1')).toBe(true);
    expect(service.tryAcquire('u1')).toBe(false);

    service.release('u1');

    expect(service.tryAcquire('u1')).toBe(true);
  });

  it('never goes negative when release is called more times than acquire (defensive against double-release)', () => {
    const service = serviceWithCap(1);

    service.release('u1');
    service.release('u1');

    expect(service.currentCount('u1')).toBe(0);
    expect(service.tryAcquire('u1')).toBe(true);
  });

  it('does not throw releasing a user that was never acquired', () => {
    const service = serviceWithCap(2);
    expect(() => service.release('never-seen')).not.toThrow();
  });
});
