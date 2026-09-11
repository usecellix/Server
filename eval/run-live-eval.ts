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
 *   1. Start the backend with a real OPENROUTER_API_KEY configured, and (since
 *      ConversationController is behind AuthGuard — TASKS.md #164)
 *      CELLIX_EVAL_BYPASS_TOKEN set to any value, with NODE_ENV != production:
 *        CELLIX_EVAL_BYPASS_TOKEN=local-eval npm run start:dev
 *   2. In another terminal, with the SAME token:
 *        CELLIX_EVAL_BYPASS_TOKEN=local-eval npm run eval:live
 *      Optionally target a non-default port/host:
 *        CELLIX_EVAL_BASE_URL=http://localhost:4001 CELLIX_EVAL_BYPASS_TOKEN=local-eval npm run eval:live
 *      Omitting CELLIX_EVAL_BYPASS_TOKEN here sends no auth header at all and
 *      every case will fail with HTTP 401 — see auth.guard.ts.
 *
 * Scoring is intentionally coarse — "did the expected action types appear /
 * did the forbidden ones not appear" — because exact-output matching against a
 * non-deterministic model is a losing game. This catches the failure mode this
 * session was built around: the model silently doing less (or something
 * different) than asked, not stylistic differences in HOW it did it.
 */
import { LIVE_GOLDEN_SET, LiveGoldenCase } from './golden-set';

const BASE_URL = process.env.CELLIX_EVAL_BASE_URL ?? 'http://localhost:4001';
const EVAL_BYPASS_TOKEN = process.env.CELLIX_EVAL_BYPASS_TOKEN;
const EVAL_BYPASS_HEADER = 'x-cellix-eval-bypass';

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
  // Bug found running this harness live for the model-swap eval: `sheetData`
  // is what SheetAnalyzerService.analyze() actually reads to detect the
  // header row and build the Planner/Executor's real cell `values` grid
  // (workbook-context.builder.ts's buildSheetContext) — the rich
  // `workbookContext.sheets[].headers`/`sampleData` sent below only supplies
  // metadata (usedRange, structure hints), it is NOT where the model's cell
  // data comes from. Sending sheetRows (data only, no header row) as
  // `sheetData` made analyze() see a data row where it expects headers, fail
  // header detection, and silently hand the model an empty A1:A1 sheet —
  // every case "ran" against a blank workbook regardless of goldenCase's real
  // fixture. `sheetData` must match what the real add-in sends: headers as
  // row 0, data rows after.
  const sheetDataWithHeaderRow = [goldenCase.sheetHeaders, ...goldenCase.sheetRows];
  const body = {
    message: goldenCase.prompt,
    sheetData: sheetDataWithHeaderRow,
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
      headers: {
        'Content-Type': 'application/json',
        ...(EVAL_BYPASS_TOKEN ? { [EVAL_BYPASS_HEADER]: EVAL_BYPASS_TOKEN } : {}),
      },
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
    const authHint =
      response.status === 401
        ? ' — set CELLIX_EVAL_BYPASS_TOKEN to the same value on both the backend and this script (see this file\'s header comment)'
        : '';
    return {
      id: goldenCase.id,
      category: goldenCase.category,
      passed: false,
      detail: `HTTP ${response.status} — is the backend running at ${BASE_URL} with an API key configured?${authHint}`,
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
  // Bug found adding the router-ambiguous cases: this used to fire
  // unconditionally, so any case correctly expecting zero actions (a
  // read-only answer, or a routing probe with several valid action shapes)
  // always failed regardless of actual correctness — live-data-query-sum was
  // silently affected before allowZeroActions existed.
  if (actions.length === 0 && !goldenCase.allowZeroActions) {
    failures.push('no actions emitted at all');
  }

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
  // Comma-separated golden-case ids to skip this run — for isolating a case
  // that's hanging/erroring without editing golden-set.ts.
  const skipIds = new Set(
    (process.env.CELLIX_EVAL_SKIP_IDS ?? '')
      .split(',')
      .map((id) => id.trim())
      .filter(Boolean),
  );
  const cases = LIVE_GOLDEN_SET.filter((c) => !skipIds.has(c.id));
  if (skipIds.size > 0) {
    console.log(`Skipping ${skipIds.size} case(s) via CELLIX_EVAL_SKIP_IDS: ${[...skipIds].join(', ')}`);
  }
  console.log(`Running ${cases.length} live golden-set case(s) against ${BASE_URL}...\n`);

  const outcomes: CaseOutcome[] = [];
  for (const goldenCase of cases) {
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
