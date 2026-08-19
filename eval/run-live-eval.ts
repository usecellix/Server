/**
 * The live half of the golden-set eval (see eval/golden-set.ts for the design).
 *
 * Unlike test/golden-set-eval.spec.ts (deterministic, free, runs in CI), this
 * sends real prompts through the real HTTP endpoint to a running backend —
 * exercising the actual LLM Router → Tier 0-3 pipeline → Executor, exactly as
 * production does. That means it costs real OpenRouter credits and is not
 * fully deterministic (models vary run to run), so it is NOT part of `npm test`
 * and must be run deliberately.
 *
 * Usage:
 *   1. Start the backend with a real OPENROUTER_API_KEY configured:
 *        npm run start:dev
 *   2. In another terminal:
 *        npm run eval:live
 *      Optionally target a non-default port/host:
 *        CELLIX_EVAL_BASE_URL=http://localhost:4001 npm run eval:live
 *
 * Scoring is intentionally coarse — "did the expected action types appear /
 * did the forbidden ones not appear" — because exact-output matching against a
 * non-deterministic model is a losing game. This catches the failure mode this
 * session was built around: the model silently doing less (or something
 * different) than asked, not stylistic differences in HOW it did it.
 */
import { LIVE_GOLDEN_SET, LiveGoldenCase } from './golden-set';

const BASE_URL = process.env.CELLIX_EVAL_BASE_URL ?? 'http://localhost:4001';

interface SheetActionLike {
  type: string;
  [key: string]: unknown;
}

interface CaseOutcome {
  id: string;
  category: string;
  passed: boolean;
  detail: string;
  actionTypesSeen: string[];
}

async function runCase(goldenCase: LiveGoldenCase): Promise<CaseOutcome> {
  const body = {
    message: goldenCase.prompt,
    sheetData: goldenCase.sheetRows,
    workbookContext: {
      sheets: [
        {
          sheetName: 'Sheet1',
          usedRange: `A1:${String.fromCharCode(64 + goldenCase.sheetHeaders.length)}${goldenCase.sheetRows.length + 1}`,
          rowCount: goldenCase.sheetRows.length + 1,
          colCount: goldenCase.sheetHeaders.length,
          headers: goldenCase.sheetHeaders,
          sampleData: goldenCase.sheetRows,
        },
      ],
      activeSheet: 'Sheet1',
    },
  };

  let response: Response;
  try {
    response = await fetch(`${BASE_URL}/excel-ai/conversation`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err);
    return {
      id: goldenCase.id,
      category: goldenCase.category,
      passed: false,
      detail: `Could not reach ${BASE_URL} (${reason}) — is the backend running? Try: npm run start:dev`,
      actionTypesSeen: [],
    };
  }

  if (!response.ok || !response.body) {
    return {
      id: goldenCase.id,
      category: goldenCase.category,
      passed: false,
      detail: `HTTP ${response.status} — is the backend running at ${BASE_URL} with an API key configured?`,
      actionTypesSeen: [],
    };
  }

  const actions = await collectActionsFromSse(response.body);
  const actionTypesSeen = actions.map((a) => a.type);

  const missing = goldenCase.mustIncludeActionTypes.filter(
    (type) => !actionTypesSeen.includes(type),
  );
  const forbidden = (goldenCase.mustNotIncludeActionTypes ?? []).filter((type) =>
    actionTypesSeen.includes(type),
  );

  const failures: string[] = [];
  if (missing.length > 0) failures.push(`missing required action type(s): ${missing.join(', ')}`);
  if (forbidden.length > 0) failures.push(`emitted forbidden action type(s): ${forbidden.join(', ')}`);
  if (actions.length === 0) failures.push('no actions emitted at all');

  return {
    id: goldenCase.id,
    category: goldenCase.category,
    passed: failures.length === 0,
    detail: failures.join('; '),
    actionTypesSeen,
  };
}

/** Minimal SSE parser: reads `event:`/`data:` lines, collects every `actions` event's payload. */
async function collectActionsFromSse(
  body: ReadableStream<Uint8Array>,
): Promise<SheetActionLike[]> {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  const collected: SheetActionLike[] = [];

  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    const frames = buffer.split('\n\n');
    buffer = frames.pop() ?? '';

    for (const frame of frames) {
      const eventLine = frame.split('\n').find((l) => l.startsWith('event:'));
      const dataLine = frame.split('\n').find((l) => l.startsWith('data:'));
      if (!eventLine || !dataLine) continue;
      const eventName = eventLine.slice('event:'.length).trim();
      if (eventName !== 'actions') continue;
      try {
        const payload = JSON.parse(dataLine.slice('data:'.length).trim()) as {
          actions?: SheetActionLike[];
        };
        collected.push(...(payload.actions ?? []));
      } catch {
        // Malformed frame — ignore, the case will fail on missing action types instead.
      }
    }
  }

  return collected;
}

async function main(): Promise<void> {
  console.log(`Running ${LIVE_GOLDEN_SET.length} live golden-set case(s) against ${BASE_URL}...\n`);

  const outcomes: CaseOutcome[] = [];
  for (const goldenCase of LIVE_GOLDEN_SET) {
    process.stdout.write(`  ${goldenCase.id}... `);
    const outcome = await runCase(goldenCase);
    outcomes.push(outcome);
    console.log(outcome.passed ? 'PASS' : `FAIL (${outcome.detail})`);
  }

  const passed = outcomes.filter((o) => o.passed).length;
  console.log(`\n=== Live golden-set: ${passed}/${outcomes.length} passed ===`);
  for (const o of outcomes.filter((o) => !o.passed)) {
    console.log(`  FAIL ${o.id} [${o.category}]: ${o.detail}`);
    console.log(`    action types seen: ${o.actionTypesSeen.join(', ') || '(none)'}`);
  }

  process.exit(passed === outcomes.length ? 0 : 1);
}

main().catch((err) => {
  console.error('Live eval crashed:', err);
  process.exit(1);
});
