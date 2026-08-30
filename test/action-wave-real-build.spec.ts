import fs from 'fs';
import path from 'path';
import { splitIntoActionWaves } from '../src/excel-ai/utils/action-wave.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #160 — staging verified against a REAL captured build.
 *
 * The fixture is the actual 140-action list a live monthly-ledger run produced
 * (captured from the SSE stream), not a synthetic shape. That distinction is
 * the standing lesson from #143/#144/#157: every one of those bugs survived
 * thorough unit tests because the fixtures encoded what the author assumed the
 * model emits rather than what it actually emits.
 */
describe('splitIntoActionWaves — real 140-action ledger build (TASKS.md #160)', () => {
  const actions: SheetActionPayload[] = JSON.parse(
    fs.readFileSync(path.join(__dirname, '__fixtures__real-ledger-actions.json'), 'utf8'),
  );

  it('the fixture is the real thing, not a stub', () => {
    expect(actions.length).toBe(140);
    const types = new Set(actions.map((a) => a.type));
    // The shapes a real run actually produced, including BATCH_SET (#157).
    expect(types).toContain('ADD_SHEET');
    expect(types).toContain('BATCH_SET');
    expect(types).toContain('FORMAT_RANGE');
    expect(types).toContain('CREATE_CHART');
  });

  it('splits into multiple reviewable steps', () => {
    const waves = splitIntoActionWaves(actions);
    expect(waves.length).toBeGreaterThan(1);
    expect(waves.length).toBeLessThanOrEqual(6);
  });

  it('loses nothing and duplicates nothing', () => {
    const flat = splitIntoActionWaves(actions).flatMap((w) => w.actions);
    expect(flat).toHaveLength(actions.length);
    expect(new Set(flat).size).toBe(actions.length);
  });

  it('creates all 13 sheets in the first step, before any write', () => {
    const waves = splitIntoActionWaves(actions);
    const creates = actions.filter((a) => a.type === 'ADD_SHEET' || a.type === 'CREATE_SHEET');
    expect(waves[0].actions).toHaveLength(creates.length);
    expect(waves[0].actions.every((a) => a.type === 'ADD_SHEET' || a.type === 'CREATE_SHEET')).toBe(true);
    expect(waves.slice(1).flatMap((w) => w.actions).some((a) => a.type === 'ADD_SHEET')).toBe(false);
  });

  it('puts the chart in the last step, after its source data exists', () => {
    const waves = splitIntoActionWaves(actions);
    const chartStep = waves.findIndex((w) => w.actions.some((a) => a.type === 'CREATE_CHART'));
    const contentStep = waves.findIndex((w) =>
      w.actions.some((a) => a.type === 'BATCH_SET' || a.type === 'SET_CELL'),
    );
    expect(chartStep).toBeGreaterThan(contentStep);
  });

  it('applies formatting only after the content it formats', () => {
    const waves = splitIntoActionWaves(actions);
    const formatStep = waves.findIndex((w) => w.actions.some((a) => a.type === 'FORMAT_RANGE'));
    const contentStep = waves.findIndex((w) =>
      w.actions.some((a) => a.type === 'BATCH_SET' || a.type === 'SET_CELL'),
    );
    expect(formatStep).toBeGreaterThanOrEqual(contentStep);
  });

  it('autofit lands at or after formatting — it measures written content', () => {
    const waves = splitIntoActionWaves(actions);
    const autofitStep = waves.findIndex((w) => w.actions.some((a) => a.type === 'AUTOFIT_COLUMNS'));
    const contentStep = waves.findIndex((w) =>
      w.actions.some((a) => a.type === 'BATCH_SET' || a.type === 'SET_CELL'),
    );
    expect(autofitStep).toBeGreaterThan(contentStep);
  });

  it('every step carries a human label', () => {
    for (const wave of splitIntoActionWaves(actions)) {
      expect(wave.label).toBeTruthy();
      expect(wave.label).not.toMatch(/^[A-Z_]+$/); // never a raw action type
    }
  });

  it('no step is empty', () => {
    for (const wave of splitIntoActionWaves(actions)) {
      expect(wave.actions.length).toBeGreaterThan(0);
    }
  });
});
