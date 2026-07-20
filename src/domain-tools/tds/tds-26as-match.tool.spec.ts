import { tds26asMatch } from './tds-26as-match.tool';
import { expectStubNotImplemented } from '../test-utils/domain-tool-test.util';

describe('tds26asMatch', () => {
  it('has DomainTool signature and throws Not implemented (stub)', () => {
    expectStubNotImplemented(tds26asMatch, {
      books: [],
      form26as: [],
      amountTolerance: 1,
    });
  });
});
