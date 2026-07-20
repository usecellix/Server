import { trialBalanceCheck } from './trial-balance-check.tool';
import { expectStubNotImplemented } from '../test-utils/domain-tool-test.util';

describe('trialBalanceCheck', () => {
  it('has DomainTool signature and throws Not implemented (stub)', () => {
    expectStubNotImplemented(trialBalanceCheck, {
      rows: [],
      tolerance: 0.01,
    });
  });
});
