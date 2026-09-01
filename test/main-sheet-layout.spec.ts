import { applyPresentationPass, collectSheetCells } from '../src/excel-ai/utils/presentation-pass.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';
import * as fs from 'fs';
import * as path from 'path';

const realLedgerActions = JSON.parse(
  fs.readFileSync(path.join(__dirname, '__fixtures__real-ledger-actions.json'), 'utf8'),
);

/**
 * TASKS.md #164 — the Main/dashboard sheet.
 *
 * All three cases below are replayed from the SAME live 140-action run the
 * user reported as "not proper", with the run's own presentation output
 * stripped so the pass re-derives it (the #143/#144/#157 lesson: fixtures come
 * from what the model really emits, not from invented shapes).
 */

const PRESENTATION_TYPES = new Set(['FORMAT_RANGE', 'SET_ROW_HEIGHT', 'FREEZE_PANES', 'AUTOFIT_COLUMNS']);

function liveMainInput(): SheetActionPayload[] {
  const raw = realLedgerActions as unknown as SheetActionPayload[] | { actions: SheetActionPayload[] };
  const list = Array.isArray(raw) ? raw : raw.actions;
  return list.filter((a) => !PRESENTATION_TYPES.has(String(a.type)));
}

function mainOutput(): SheetActionPayload[] {
  const out = applyPresentationPass(liveMainInput(), { userMessage: 'monthly ledger with a main dashboard' });
  return out.filter((a) => String(a.sheetName ?? '').toLowerCase() === 'main');
}

function formatsAt(out: SheetActionPayload[], row: number, col: number): SheetActionPayload[] {
  return out.filter((a) => a.type === 'FORMAT_RANGE' && a.row === row && a.col === col);
}

describe('collectSheetCells reads every action shape', () => {
  it('sees a title written inside a BATCH_SET, not only SET_CELL', () => {
    const cells = collectSheetCells(
      [
        {
          type: 'BATCH_SET',
          sheetName: 'Main',
          operations: [
            { address: 'A1', value: 'Dashboard' },
            { address: 'B2', formula: '=SUM(B5:B16)' },
          ],
        } as SheetActionPayload,
      ],
      'Main',
    );
    expect(cells).toEqual([
      { row: 0, col: 0, label: 'Dashboard', isValue: false },
      { row: 1, col: 1, label: null, isValue: true },
    ]);
  });

  it('ignores other sheets', () => {
    const actions = [
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0, value: 'x' },
    ] as SheetActionPayload[];
    expect(collectSheetCells(actions, 'Main')).toEqual([]);
  });
});

describe('#164a — the dashboard title is styled again', () => {
  it('styles A1 "Dashboard" even though it arrives in a BATCH_SET', () => {
    const title = formatsAt(mainOutput(), 0, 0);
    expect(title).toHaveLength(1);
    expect(title[0].format).toMatchObject({ bold: true, fontSize: 14 });
  });
});

describe('#164b — the KPI band is styled', () => {
  it('gives every label/value tile a fill and a border', () => {
    const out = mainOutput();
    // Live band: A2/B2, C2/D2, E2/F2 (row index 1).
    for (const col of [0, 1, 2, 3, 4, 5]) {
      const tiles = formatsAt(out, 1, col);
      expect(tiles).toHaveLength(1);
      expect(tiles[0].format).toMatchObject({ fillColor: '#EDF3FA', borders: 'all' });
    }
  });

  it('currency-formats the VALUE cells, not the label cells', () => {
    const out = mainOutput();
    expect(formatsAt(out, 1, 1)[0].format).toHaveProperty('numberFormat');
    expect(formatsAt(out, 1, 3)[0].format).toHaveProperty('numberFormat');
    expect(formatsAt(out, 1, 5)[0].format).toHaveProperty('numberFormat');
    expect(formatsAt(out, 1, 0)[0].format).not.toHaveProperty('numberFormat');
    expect(formatsAt(out, 1, 2)[0].format).not.toHaveProperty('numberFormat');
  });

  it('does not treat the KPI row as a table header', () => {
    // #144: no header band (white-on-blue) may land on the KPI row.
    const banded = mainOutput().filter(
      (a) => a.type === 'FORMAT_RANGE' && a.row === 1 && (a.format as never as { fillColor?: string })?.fillColor === '#2F5597',
    );
    expect(banded).toHaveLength(0);
  });

  it('needs at least two pairs before styling anything as a band', () => {
    const out = applyPresentationPass(
      [
        { type: 'ADD_SHEET', sheetName: 'Solo' } as SheetActionPayload,
        {
          type: 'BATCH_SET',
          sheetName: 'Solo',
          operations: [
            { address: 'A1', value: 'Total' },
            { address: 'B1', formula: '=SUM(B3:B9)' },
            { address: 'A2', value: 'Month' },
            { address: 'B2', value: 'Amount' },
            { address: 'C2', value: 'Paid' },
          ],
        } as SheetActionPayload,
      ],
      {},
    );
    const tiles = out.filter(
      (a) => a.type === 'FORMAT_RANGE' && a.row === 0 && (a.format as never as { fillColor?: string })?.fillColor === '#EDF3FA',
    );
    expect(tiles).toHaveLength(0);
  });
});

describe('#164c — a table format runway stops at the next table', () => {
  it('does not paint the summary table currency over the consolidated table below', () => {
    const out = mainOutput();
    // Summary header is row 3 (A4); consolidated header is row 17 (A18).
    for (const col of [1, 2, 3]) {
      const runway = formatsAt(out, 4, col);
      expect(runway).toHaveLength(1);
      // Rows 4..16 inclusive — never reaching row 17.
      expect(runway[0].rowCount).toBe(13);
      expect(4 + (runway[0].rowCount as number)).toBeLessThanOrEqual(17);
    }
  });

  it('keeps the full runway for the LAST table on a sheet', () => {
    const consolidated = mainOutput().filter((a) => a.type === 'FORMAT_RANGE' && a.row === 18);
    expect(consolidated.length).toBeGreaterThan(0);
    for (const action of consolidated) expect(action.rowCount).toBe(200);
  });

  it('leaves a single-table sheet unchanged — every month sheet keeps 200', () => {
    const out = applyPresentationPass(liveMainInput(), { userMessage: 'monthly ledger' });
    const january = out.filter(
      (a) => String(a.sheetName) === 'January' && a.type === 'FORMAT_RANGE' && a.row === 1,
    );
    expect(january.length).toBeGreaterThan(0);
    for (const action of january) expect(action.rowCount).toBe(200);
  });
});
