/**
 * Ad-hoc single-prompt probe against a running backend, using the same fixture
 * as `usecase-probe.ts`. Handy when checking whether a phrasing change moves an
 * answer (e.g. suspecting a cached response).
 *
 *   MESSAGE="Describe this sheet" npx ts-node eval/one-off-probe.ts
 */
import { buildFixture } from './usecase-fixture';

const BASE_URL = process.env.CELLIX_EVAL_BASE_URL ?? 'http://localhost:4001';
const TOKEN = process.env.CELLIX_EVAL_BYPASS_TOKEN ?? 'usecase-probe';
const MESSAGE = process.env.MESSAGE ?? 'Describe this spreadsheet to me';

async function main(): Promise<void> {
  const fixture = buildFixture();
  const response = await fetch(`${BASE_URL}/excel-ai/conversation`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-cellix-eval-bypass': TOKEN },
    body: JSON.stringify({
      message: MESSAGE,
      sheetData: fixture.sheetData,
      workbookContext: fixture.workbookContext,
      mode: 'action',
    }),
  });

  const raw = await response.text();
  const events = raw
    .split('\n\n')
    .map((frame) => {
      const event = frame.split('\n').find((line) => line.startsWith('event:'))?.slice(6).trim();
      const data = frame
        .split('\n')
        .filter((line) => line.startsWith('data:'))
        .map((line) => line.slice(5).trim())
        .join('');
      return { event, data };
    })
    .filter((entry) => entry.event && ['answer', 'chunk', 'actions', 'error'].includes(entry.event));

  let text = '';
  for (const { event, data } of events) {
    try {
      const parsed = JSON.parse(data);
      if (event === 'chunk') text += parsed.text ?? '';
      else text += `\n[${event}] ${JSON.stringify(parsed).slice(0, 1500)}`;
    } catch {
      text += `\n[${event}] ${data.slice(0, 300)}`;
    }
  }
  console.log(text.trim());
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
