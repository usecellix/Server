import { parseAgentJson, parseExecutorPayload } from '../src/agents/utils/parse-agent-json.util';

describe('parseAgentJson', () => {
  it('parses planner payload wrapped in markdown fences', () => {
    const raw = '```json\n{"subtasks":[{"id":"s1","description":"Do work","targetSheet":"Sheet1"}]}\n```';
    const parsed = parseAgentJson<{ subtasks: Array<{ id: string }> }>(raw);
    expect(parsed.subtasks[0].id).toBe('s1');
  });

  it('parses planner payload embedded in prose', () => {
    const raw = 'Here is the output you asked for:\n{"subtasks":[],"clarificationsNeeded":[]}\nThanks!';
    const parsed = parseAgentJson<{ clarificationsNeeded: string[] }>(raw);
    expect(parsed.clarificationsNeeded).toEqual([]);
  });
});

describe('parseExecutorPayload', () => {
  it('parses fenced executor payload', () => {
    const raw = '```json\n{"subtaskId":"s1","actions":[],"isDone":true}\n```';
    expect(parseExecutorPayload(raw)).toEqual({
      subtaskId: 's1',
      actions: [],
      isDone: true,
    });
  });

  it('wraps a bare action array as executor payload', () => {
    const raw = 'Executor output:\n[{"type":"ADD_ROW","data":["GST","","=C10*0.1"]}]';
    expect(parseExecutorPayload(raw)).toEqual({
      actions: [{ type: 'ADD_ROW', data: ['GST', '', '=C10*0.1'] }],
      isDone: true,
    });
  });
});
