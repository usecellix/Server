import { WriteRouteNoActionError } from '../src/excel-ai/errors/write-route-no-action.error';

describe('WriteRouteNoActionError', () => {
  it('exposes a distinct user-facing message and error code', () => {
    const error = new WriteRouteNoActionError('conv_1', 'sort the sheet');

    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('WriteRouteNoActionError');
    expect(error.code).toBe('WRITE_ROUTE_NO_ACTION');
    expect(error.conversationId).toBe('conv_1');
    expect(error.userMessage).toBe('sort the sheet');
    expect(error.message).toBe(
      'Something went wrong applying this change — try rephrasing',
    );
  });
});
