import { costAllocation } from './cost-allocation.tool';
import { expectStubNotImplemented } from '../test-utils/domain-tool-test.util';

describe('costAllocation', () => {
  it('has DomainTool signature and throws Not implemented (stub)', () => {
    expectStubNotImplemented(costAllocation, {
      costPool: 100000,
      bases: [],
    });
  });
});
