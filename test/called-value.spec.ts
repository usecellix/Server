import {
  extractCalledValueFromMessage,
  resolveRemarkValue,
  stripCalledLabel,
} from '../src/agents/utils/called-value.util';
import {
  annotateClearIntentOverwrite,
  isClearOrEmptyIntent,
} from '../src/agents/utils/clear-intent-overwrite.util';
import {
  normalizeSingleAction,
  normalizeExecutorOutput,
} from '../src/agents/utils/normalize-executor-output.util';

describe('called-value util', () => {
  it('extracts Cleared from "Called Cleared" phrasing', () => {
    expect(
      extractCalledValueFromMessage('Add remarks to paid invoices , Called Cleared'),
    ).toBe('Cleared');
    expect(extractCalledValueFromMessage('mark paid rows called Follow-up')).toBe(
      'Follow-up',
    );
    expect(extractCalledValueFromMessage('set status, call it Done')).toBe('Done');
  });

  it('strips a Called prefix from cell values', () => {
    expect(stripCalledLabel('Called Cleared')).toBe('Cleared');
    expect(stripCalledLabel('Cleared')).toBe('Cleared');
  });

  it('resolves model value using the user message', () => {
    expect(
      resolveRemarkValue('Called Cleared', 'Add remarks to paid invoices, Called Cleared'),
    ).toBe('Cleared');
  });
});

describe('clear-intent overwrite', () => {
  it('detects empty/clear remarks prompts', () => {
    expect(isClearOrEmptyIntent('make the remarks column empty no values')).toBe(true);
    expect(isClearOrEmptyIntent('clear the Remarks column')).toBe(true);
    expect(isClearOrEmptyIntent('add remarks Cleared')).toBe(false);
  });

  it('marks blank SET_MATCHING_ROWS as overwrite-confirmed', () => {
    const [action] = annotateClearIntentOverwrite(
      [
        {
          type: 'SET_MATCHING_ROWS',
          sheetName: 'Purchase Register',
          range: 'A1:L51',
          hasHeaders: true,
          targetColumn: 'Remarks',
          value: '',
        },
      ],
      'make the remarks column empty no values',
    );
    expect((action as { explicitOverwriteConfirmed?: boolean }).explicitOverwriteConfirmed).toBe(
      true,
    );
  });
});

describe('SET_MATCHING_ROWS normalization', () => {
  it('accepts SET_MATCHING_ROWS and strips Called from the value', () => {
    const action = normalizeSingleAction(
      {
        type: 'SET_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A1:L51',
        hasHeaders: true,
        filter: { column: 'Payment Status', operator: 'equals', value: 'Paid' },
        targetColumn: 'Remarks',
        value: 'Called Cleared',
      },
      'Purchase Register',
    );
    expect(action?.type).toBe('SET_MATCHING_ROWS');
    expect(action?.targetColumn).toBe('Remarks');
    expect(action?.value).toBe('Cleared');
  });

  it('applies called-value fix from subtask description on executor output', () => {
    const output = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'SET_MATCHING_ROWS',
            sheetName: 'Purchase Register',
            range: 'A1:L51',
            filter: { column: 'Payment Status', operator: 'equals', value: 'Paid' },
            targetColumn: 'Remarks',
            value: 'Called Cleared',
          },
        ],
        isDone: true,
      },
      {
        id: 's1',
        description: 'Add remarks to paid invoices, Called Cleared',
        targetSheet: 'Purchase Register',
        dependsOn: [],
        estimatedActions: 1,
        suggestedActionType: 'SET_MATCHING_ROWS',
      },
    );
    expect(output.actions[0]?.type).toBe('SET_MATCHING_ROWS');
    expect(output.actions[0]?.value).toBe('Cleared');
  });

  it('confirms overwrite for clear-empty SET_MATCHING_ROWS from subtask text', () => {
    const output = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'SET_MATCHING_ROWS',
            sheetName: 'Purchase Register',
            range: 'A1:L51',
            targetColumn: 'Remarks',
            value: '',
          },
        ],
        isDone: true,
      },
      {
        id: 's1',
        description: 'Make the Remarks column empty on Purchase Register',
        targetSheet: 'Purchase Register',
        dependsOn: [],
        estimatedActions: 1,
        suggestedActionType: 'SET_MATCHING_ROWS',
      },
    );
    expect(output.actions[0]?.explicitOverwriteConfirmed).toBe(true);
    expect(output.actions[0]?.value).toBe('');
  });
});
