import { LlmRouterService } from '../src/excel-ai/services/llm-router.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { SheetAnalyzerService } from '../src/excel-ai/services/sheet-analyzer.service';
import {
  ROUTING_GOLDEN_SET,
  HEADER_DETECTION_GOLDEN_SET,
  RoutingGoldenCase,
} from '../eval/golden-set';

/**
 * The deterministic half of the golden-set eval (see eval/golden-set.ts for the
 * full design). This is the project's reliability scorecard for the parts of
 * the pipeline that don't require a real LLM call — routing/tier classification
 * and header detection. It runs in CI on every change, for free, and is meant
 * to fail loudly the moment a routing regression like the purchase-register bug
 * ships again.
 *
 * What this does NOT cover: whether the LLM's actual output (the Planner's
 * subtasks, the Executor's actions) is correct. That requires a real API call —
 * see eval/run-live-eval.ts for the opt-in, costed counterpart.
 */

/** Fails the test with a clear message if the router calls out to the LLM. */
function unreachableOpenRouter(): OpenRouterService {
  return {
    complete: jest.fn(async () => {
      throw new Error(
        'This golden case is expected to resolve deterministically (regex/complexity ' +
          'classifier/write-intent-guard) — it should never reach the LLM Router. If this ' +
          'prompt legitimately needs an LLM call now, move it out of the deterministic ' +
          'golden set rather than mocking a specific response here.',
      );
    }),
  } as unknown as OpenRouterService;
}

interface CaseResult {
  id: string;
  category: string;
  passed: boolean;
  detail: string;
}

describe('Golden-set eval — deterministic routing scorecard', () => {
  const router = new LlmRouterService(unreachableOpenRouter(), {} as AppConfigService);
  const results: CaseResult[] = [];

  afterAll(() => {
    printScorecard('Routing', results);
  });

  it.each(ROUTING_GOLDEN_SET)('$id — $prompt', async (goldenCase: RoutingGoldenCase) => {
    const decision = await router.route({
      message: goldenCase.prompt,
      mode: goldenCase.mode ?? 'action',
      sheetHeaders: goldenCase.sheetHeaders ?? [],
      activeSheet: 'Sheet1',
    });

    const failures: string[] = [];
    if (decision.route !== goldenCase.expected.route) {
      failures.push(`route: expected "${goldenCase.expected.route}", got "${decision.route}"`);
    }
    if (goldenCase.expected.tier !== undefined && decision.complexity !== goldenCase.expected.tier) {
      failures.push(`tier: expected ${goldenCase.expected.tier}, got ${decision.complexity}`);
    }
    if (
      goldenCase.expected.actionHint !== undefined &&
      decision.actionHint !== goldenCase.expected.actionHint
    ) {
      failures.push(
        `actionHint: expected "${goldenCase.expected.actionHint}", got "${decision.actionHint}"`,
      );
    }

    results.push({
      id: goldenCase.id,
      category: goldenCase.category,
      passed: failures.length === 0,
      detail: failures.join('; '),
    });

    if (failures.length > 0) {
      throw new Error(failures.join('; '));
    }
  });
});

describe('Golden-set eval — header detection scorecard', () => {
  const analyzer = new SheetAnalyzerService();
  const results: CaseResult[] = [];

  afterAll(() => {
    printScorecard('Header detection', results);
  });

  it.each(HEADER_DETECTION_GOLDEN_SET)('$id — $description', (goldenCase) => {
    const analysis = analyzer.analyze(goldenCase.sheetData);

    const failures: string[] = [];
    if (analysis.headerRowIndex !== goldenCase.expectedHeaderRowIndex) {
      failures.push(
        `headerRowIndex: expected ${goldenCase.expectedHeaderRowIndex}, got ${analysis.headerRowIndex}`,
      );
    }
    if (JSON.stringify(analysis.headers) !== JSON.stringify(goldenCase.expectedHeaders)) {
      failures.push(
        `headers: expected ${JSON.stringify(goldenCase.expectedHeaders)}, got ${JSON.stringify(analysis.headers)}`,
      );
    }

    results.push({
      id: goldenCase.id,
      category: goldenCase.category,
      passed: failures.length === 0,
      detail: failures.join('; '),
    });

    if (failures.length > 0) {
      throw new Error(failures.join('; '));
    }
  });
});

function printScorecard(label: string, results: CaseResult[]): void {
  if (results.length === 0) return;
  const passed = results.filter((r) => r.passed).length;
  const byCategory = new Map<string, { passed: number; total: number }>();
  for (const r of results) {
    const bucket = byCategory.get(r.category) ?? { passed: 0, total: 0 };
    bucket.total += 1;
    if (r.passed) bucket.passed += 1;
    byCategory.set(r.category, bucket);
  }

  const lines = [
    `\n=== ${label} golden-set scorecard: ${passed}/${results.length} (${((passed / results.length) * 100).toFixed(1)}%) ===`,
    ...[...byCategory.entries()].map(
      ([cat, { passed: p, total }]) => `  ${cat}: ${p}/${total}`,
    ),
    ...results.filter((r) => !r.passed).map((r) => `  FAIL ${r.id}: ${r.detail}`),
  ];
  // eslint-disable-next-line no-console
  console.log(lines.join('\n'));
}
