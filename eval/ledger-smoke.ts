/**
 * End-to-end smoke test for the 12-month booking-ledger prompt — the live
 * failure behind TASKS.md #259-#263.
 *
 * Drives the REAL backend the way the task pane does: one POST to
 * /excel-ai/conversation, then a POST to /conversation/continue per
 * `wave_ready`, accepting every wave, until the run ends. Every action the run
 * emits is replayed against a virtual workbook so the two things that actually
 * broke in Excel can be asserted without Excel:
 *
 *   1. a write landing on a sheet no action ever created ("The requested
 *      resource doesn't exist" at Accept — what made Accept look dead), and
 *   2. a formula referencing a sheet that never gets created (#REF!).
 *
 *   CELLIX_EVAL_BASE_URL=http://localhost:4011 npx ts-node eval/ledger-smoke.ts
 */

const BASE_URL = process.env.CELLIX_EVAL_BASE_URL ?? 'http://localhost:4011';
const TOKEN = process.env.CELLIX_EVAL_BYPASS_TOKEN ?? 'usecase-probe';
const MAX_WAVES = Number(process.env.MAX_WAVES ?? 25);

const PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the ' +
  'details of the remaining sheets, in the main sheet i need to have dashboard also, my need to ' +
  'record payments and related things ,which all month sheets include Unit No, Guest, Guest name, ' +
  'check in, check out, Rate per night, total amount, source, payment status, bank account';

/**
 * An empty workbook — exactly what the user ran against.
 *
 * Shape matches `usecase-fixture.ts` exactly: the add-in sends `sheetName`,
 * not `name`. Getting that wrong is what surfaced the crash fixed in #263.
 */
function emptyWorkbook() {
  return {
    sheetData: [[]],
    workbookContext: {
      activeSheet: 'Sheet1',
      sheets: [
        {
          sheetName: 'Sheet1',
          usedRange: 'A1',
          rowCount: 0,
          colCount: 0,
          headers: [] as string[],
          sampleData: [] as unknown[][],
          isHidden: false,
        },
      ],
    },
  };
}

interface SseEvent {
  event: string;
  data: Record<string, unknown>;
}

function parseSse(raw: string): SseEvent[] {
  const out: SseEvent[] = [];
  for (const frame of raw.split('\n\n')) {
    const event = frame
      .split('\n')
      .find((l) => l.startsWith('event:'))
      ?.slice(6)
      .trim();
    const data = frame
      .split('\n')
      .filter((l) => l.startsWith('data:'))
      .map((l) => l.slice(5).trim())
      .join('');
    if (!event || !data) continue;
    try {
      out.push({ event, data: JSON.parse(data) as Record<string, unknown> });
    } catch {
      /* non-JSON frame — ignore */
    }
  }
  return out;
}

type Action = Record<string, unknown> & { type: string };

/** Sheet names an action creates. */
function createdSheet(a: Action): string | null {
  if (a.type !== 'ADD_SHEET' && a.type !== 'CREATE_SHEET') return null;
  const name = String(a.name ?? a.sheetName ?? '').trim();
  return name || null;
}

/** Sheet an action writes to, when it needs that sheet to already exist. */
function writesToSheet(a: Action): string | null {
  if (createdSheet(a)) return null;
  const structural = new Set(['DELETE_SHEET', 'RENAME_SHEET', 'COPY_SHEET', 'MOVE_SHEET']);
  if (structural.has(a.type)) return null;
  const name = String(a.sheetName ?? a.destSheet ?? '').trim();
  return name || null;
}

const SHEET_REF = /(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!\$?[A-Za-z]{1,3}\$?\d*/g;

/** Sheets a formula in this action reads from. */
function referencedSheets(a: Action): string[] {
  const texts: string[] = [];
  const collect = (v: unknown) => {
    if (typeof v === 'string' && v.startsWith('=')) texts.push(v);
  };
  collect(a.formula);
  if (Array.isArray(a.operations)) {
    for (const op of a.operations as Array<Record<string, unknown>>) {
      collect(op.formula);
      collect(op.value);
    }
  }
  const names = new Set<string>();
  for (const text of texts) {
    for (const m of text.matchAll(SHEET_REF)) {
      const n = (m[1]?.replace(/''/g, "'") ?? m[2] ?? '').trim();
      if (n) names.add(n);
    }
  }
  return [...names];
}

/**
 * Deliberately `node:http` rather than `fetch`. Node's fetch (undici) applies
 * a 5-minute body timeout that cannot be configured without the `undici`
 * package, and a legitimate wave here can run to the server's own 480s budget:
 * a real 12-month run died with `UND_ERR_BODY_TIMEOUT` mid-wave, so the
 * harness was unable to measure the very case it exists to measure.
 */
async function post(path: string, body: unknown): Promise<SseEvent[]> {
  const http = await import('node:http');
  const payload = JSON.stringify(body);
  const url = new URL(`${BASE_URL}${path}`);

  return new Promise<SseEvent[]>((resolve, reject) => {
    const req = http.request(
      {
        hostname: url.hostname,
        port: url.port,
        path: url.pathname + url.search,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(payload),
          'x-cellix-eval-bypass': TOKEN,
        },
      },
      (res) => {
        let text = '';
        res.setEncoding('utf8');
        res.on('data', (chunk: string) => {
          text += chunk;
        });
        res.on('end', () => {
          if ((res.statusCode ?? 0) < 200 || (res.statusCode ?? 0) >= 300) {
            reject(new Error(`${path} -> HTTP ${res.statusCode}: ${text.slice(0, 400)}`));
            return;
          }
          resolve(parseSse(text));
        });
      },
    );
    // No socket/response timeout: the server owns the budget, not the client.
    req.setTimeout(0);
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

async function main(): Promise<void> {
  const started = Date.now();
  const fixture = emptyWorkbook();

  // Unique per run — omitting workbookId let repeated invocations resolve to
  // the same underlying conversation record, which could carry stale history.
  const workbookId = `wb_smoke_${Date.now()}`;
  console.log(`[smoke] POST /excel-ai/conversation  (${BASE_URL})`);
  let events = await post('/excel-ai/conversation', {
    message: PROMPT,
    sheetData: fixture.sheetData,
    workbookContext: fixture.workbookContext,
    mode: 'action',
    workbookId,
  });

  const allActions: Action[] = [];
  const waveSizes: number[] = [];
  let runId: string | undefined;
  let clarification: string | undefined;
  let waves = 0;

  const absorb = (evts: SseEvent[]) => {
    for (const { event, data } of evts) {
      if (event === 'actions' && Array.isArray(data.actions)) {
        const batch = data.actions as Action[];
        allActions.push(...batch);
        waveSizes.push(batch.length);
      }
      if (event === 'clarification' && typeof data.question === 'string') {
        clarification = data.question;
      }
      if (event === 'wave_ready' && typeof data.runId === 'string') {
        runId = data.runId;
      }
      if (event === 'error') {
        console.log(`[smoke] ERROR event: ${JSON.stringify(data).slice(0, 300)}`);
      }
    }
  };

  absorb(events);

  while (
    runId &&
    events.some((e) => e.event === 'wave_ready' && e.data.hasMore !== false) &&
    waves < MAX_WAVES
  ) {
    waves += 1;
    const readyMeta = events.find((e) => e.event === 'wave_ready')?.data;
    process.stdout.write(
      `[smoke] wave ${waves} → accept … (waveIndex=${readyMeta?.waveIndex}, waveTotal=${readyMeta?.waveTotal}, hasMore=${readyMeta?.hasMore})\n`,
    );
    events = await post('/excel-ai/conversation/continue', { runId, decision: 'accepted' });
    absorb(events);
  }
  const finalMeta = events.find((e) => e.event === 'wave_ready')?.data;
  if (finalMeta) {
    process.stdout.write(
      `[smoke] loop ended (waveIndex=${finalMeta.waveIndex}, waveTotal=${finalMeta.waveTotal}, hasMore=${finalMeta.hasMore})\n`,
    );
  }

  // ---- Replay every action against a virtual workbook, in emission order ----
  const existing = new Set<string>(['sheet1']);
  const key = (n: string) => n.trim().toLowerCase();
  const createOrder: string[] = [];
  const writesBeforeCreate: Array<{ type: string; sheet: string }> = [];
  const refsToMissing: Array<{ sheet: string; missing: string }> = [];

  for (const action of allActions) {
    const created = createdSheet(action);
    if (created) {
      if (!existing.has(key(created))) createOrder.push(created);
      existing.add(key(created));
      continue;
    }
    const target = writesToSheet(action);
    if (target && !existing.has(key(target))) {
      writesBeforeCreate.push({ type: action.type, sheet: target });
    }
  }

  // Formula refs are judged against the FINAL sheet set: a formula may legally
  // reference a sheet a later wave creates.
  for (const action of allActions) {
    for (const ref of referencedSheets(action)) {
      if (!existing.has(key(ref))) {
        refsToMissing.push({ sheet: String(action.sheetName ?? '?'), missing: ref });
      }
    }
  }

  const MONTHS = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
  ];
  const missingMonths = MONTHS.filter((m) => !existing.has(key(m)));
  const hasMain = existing.has('main');

  const uniq = <T>(xs: T[]) => [...new Set(xs.map((x) => JSON.stringify(x)))].map((s) => JSON.parse(s) as T);

  console.log('\n──────── RESULT ────────');
  console.log(`elapsed        : ${((Date.now() - started) / 1000).toFixed(1)}s`);
  console.log(`waves accepted : ${waves}  (action counts: ${waveSizes.join(', ')})`);
  console.log(`total actions  : ${allActions.length}`);
  if (clarification) console.log(`clarification  : ${clarification}`);
  console.log(`sheets created : ${createOrder.length} → ${createOrder.join(', ') || '(none)'}`);
  console.log(`Main created   : ${hasMain ? 'YES' : 'NO  ← #262'}`);
  console.log(`months missing : ${missingMonths.length ? missingMonths.join(', ') : 'none'}`);

  const uniqueWrites = uniq(writesBeforeCreate);
  console.log(
    `writes onto a sheet that does not exist yet : ${uniqueWrites.length}` +
      (uniqueWrites.length
        ? `\n   ${uniqueWrites.map((w) => `${w.type} → ${w.sheet}`).join('\n   ')}`
        : ''),
  );

  const uniqueRefs = uniq(refsToMissing);
  console.log(
    `formulas referencing a never-created sheet  : ${uniqueRefs.length}` +
      (uniqueRefs.length
        ? `\n   ${uniqueRefs.map((r) => `${r.sheet} → ${r.missing}`).join('\n   ')}`
        : ''),
  );

  // TASKS.md #265 — a column width below the readability floor means real
  // text (header names) render as 1-2 clipped characters, reading as a
  // broken/blank sheet even though every value is correct.
  const MIN_READABLE_WIDTH = 40;
  const tooNarrow = allActions.filter(
    (a) => a.type === 'SET_COLUMN_WIDTH' && typeof a.width === 'number' && (a.width as number) < MIN_READABLE_WIDTH,
  );
  console.log(
    `column widths below ${MIN_READABLE_WIDTH}pt (unreadable)         : ${tooNarrow.length}` +
      (tooNarrow.length
        ? `\n   ${tooNarrow.map((a) => `${a.sheetName} col ${a.col}: ${a.width}pt`).join('\n   ')}`
        : ''),
  );

  const ok =
    hasMain &&
    missingMonths.length === 0 &&
    uniqueWrites.length === 0 &&
    uniqueRefs.length === 0 &&
    tooNarrow.length === 0;
  console.log(`\nVERDICT: ${ok ? 'PASS' : 'FAIL'}`);
  process.exit(ok ? 0 : 1);
}

main().catch((error) => {
  console.error('[smoke] fatal:', error);
  process.exit(2);
});
