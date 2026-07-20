import { bankRecon } from './bank-recon.tool';
import { expectStubNotImplemented } from '../test-utils/domain-tool-test.util';

describe('bankRecon', () => {
  it('has DomainTool signature and throws Not implemented (stub)', () => {
    expectStubNotImplemented(bankRecon, {
      bankStatement: [],
      books: [],
      amountTolerance: 1,
      dateWindowDays: 3,
    });
  });
});
