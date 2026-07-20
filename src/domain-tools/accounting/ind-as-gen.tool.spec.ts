import { indAsGen } from './ind-as-gen.tool';
import { expectStubNotImplemented } from '../test-utils/domain-tool-test.util';

describe('indAsGen', () => {
  it('has DomainTool signature and throws Not implemented (stub)', () => {
    expectStubNotImplemented(indAsGen, {
      trialBalanceRows: [],
      standard: 'IndAS',
    });
  });
});
