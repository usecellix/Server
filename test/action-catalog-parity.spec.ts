import { readFileSync } from 'fs';
import { join } from 'path';
import {
  ALL_SHEET_ACTION_TYPES,
  EXECUTOR_ADVERTISED_ACTION_TYPES,
  actionExclusionReason,
} from '../src/excel-ai/types/action-catalog';
import { EXECUTOR_SYSTEM_PROMPT } from '../src/agents/prompts/executor.prompt';
import { normalizeExecutorOutput } from '../src/agents/utils/normalize-executor-output.util';
import { SheetActionType } from '../src/excel-ai/types/sheet-actions.types';
import { SubTask } from '../src/agents/types/agent.types';

/**
 * The action catalog used to live in three hand-maintained places that drifted —
 * most damagingly FREEZE_PANES, which had a schema entry and a working handler
 * but was missing from the Executor prompt, so the model could never emit it.
 * These tests pin every derived copy back to the one exhaustive catalog.
 */
describe('action catalog parity', () => {
  const subtask: SubTask = {
    id: 's1',
    description: 'catalog parity probe',
    targetSheet: 'Sheet1',
    dependsOn: [],
    estimatedActions: 1,
  };

  it('parses the type union out of the source file for comparison', () => {
    expect(unionFromSource().length).toBeGreaterThan(0);
  });

  it('classifies every SheetActionType in the union', () => {
    // The catalog is typed Record<SheetActionType, …>, so tsc already enforces
    // this. Assert it against the source text too, to catch a union member that
    // was added while the catalog was loosened or cast.
    expect([...ALL_SHEET_ACTION_TYPES].sort()).toEqual(unionFromSource().sort());
  });

  it('advertises every action type to the Executor that is not deliberately withheld', () => {
    const missing = ALL_SHEET_ACTION_TYPES.filter(
      (type) => !EXECUTOR_SYSTEM_PROMPT.includes(type) && actionExclusionReason(type) === null,
    );
    expect(missing).toEqual([]);
  });

  it('gives a written reason for every withheld action type', () => {
    const withheld = ALL_SHEET_ACTION_TYPES.filter((type) => actionExclusionReason(type) !== null);
    // Guards against silently withholding a capability by leaving the reason blank.
    for (const type of withheld) {
      expect(actionExclusionReason(type)!.length).toBeGreaterThan(20);
    }
    expect(withheld.length + EXECUTOR_ADVERTISED_ACTION_TYPES.length).toBe(
      ALL_SHEET_ACTION_TYPES.length,
    );
  });

  // A type-specific field required for the action to be well-formed, beyond
  // the generic { type, sheetName } every other advertised type accepts. Keep
  // this in sync with hasRequiredFields in normalize-executor-output.util.ts —
  // this test's job is to prove every ADVERTISED type is actually acceptable,
  // not to relax what "well-formed" means for that type.
  const REQUIRED_EXTRA_FIELDS: Partial<Record<string, Record<string, unknown>>> = {
    BATCH_SET: { operations: [{ address: 'A1', value: 'x' }] },
    CONDITIONAL_FORMAT: {
      range: 'B2:B10',
      rule: { kind: 'cellValue', operator: 'greaterThan', value: 1000, format: { fillColor: '#FFC7CE' } },
    },
  };

  it('accepts every advertised action type through the normalizer', () => {
    // An advertised type that the normalizer rejects would be emitted by the
    // model and then silently discarded — the exact failure Phase A surfaced.
    const rejected = EXECUTOR_ADVERTISED_ACTION_TYPES.filter((type) => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [{ type, sheetName: 'Sheet1', ...REQUIRED_EXTRA_FIELDS[type] }],
        },
        subtask,
      );
      return result.actions.length === 0;
    });
    expect(rejected).toEqual([]);
  });

  it('does not advertise SET_ZOOM, which Office.js cannot perform', () => {
    // Excel's JS API exposes print zoom only; advertising it plans guaranteed failures.
    expect(EXECUTOR_ADVERTISED_ACTION_TYPES).not.toContain('SET_ZOOM');
    expect(actionExclusionReason('SET_ZOOM')).toMatch(/zoom/i);
  });

  /**
   * The add-in is a separate git repository, so its action union cannot be
   * imported at build time and is maintained by hand. Any type the backend can
   * emit but the add-in does not list is rejected there as an unsupported
   * action, so check the two agree whenever both are checked out together.
   */
  it('emits no action type the add-in would reject as unsupported', () => {
    const frontendUnion = frontendActionUnion();
    if (!frontendUnion) {
      console.warn('Skipping add-in catalog parity — frontend/ is not checked out alongside.');
      return;
    }
    const unknownToFrontend = ALL_SHEET_ACTION_TYPES.filter(
      (type) => !frontendUnion.includes(type),
    );
    expect(unknownToFrontend).toEqual([]);
  });
});

/** The add-in's action union, or null when the sibling repo is absent. */
function frontendActionUnion(): string[] | null {
  const path = join(__dirname, '../../frontend/src/types/sheet-actions.ts');
  let source: string;
  try {
    source = readFileSync(path, 'utf8');
  } catch {
    return null;
  }
  const block = source.match(/export type SheetActionType =([\s\S]*?);/);
  if (!block) return null;
  return [...block[1].matchAll(/'([A-Z_]+)'/g)].map((m) => m[1]);
}

/** Read the SheetActionType union straight out of its source file. */
function unionFromSource(): SheetActionType[] {
  const source = readFileSync(
    join(__dirname, '../src/excel-ai/types/sheet-actions.types.ts'),
    'utf8',
  );
  const block = source.match(/export type SheetActionType =([\s\S]*?);/);
  if (!block) throw new Error('Could not locate the SheetActionType union');
  return [...block[1].matchAll(/'([A-Z_]+)'/g)].map((m) => m[1] as SheetActionType);
}
