import { itcCompute } from './itc-compute.tool';
import { expectStubNotImplemented } from '../test-utils/domain-tool-test.util';

describe('itcCompute', () => {
  it('has DomainTool signature and throws Not implemented (stub)', () => {
    expectStubNotImplemented(itcCompute, {
      eligibleInvoices: [],
      igstRate: 0.18,
    });
  });
});
