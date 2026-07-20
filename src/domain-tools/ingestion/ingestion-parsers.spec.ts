import * as fs from 'fs';
import * as path from 'path';
import { parseGstr2b } from './gstr2b-parser';
import { parseForm26as } from './form26as-parser';
import { parseTallyExport } from './tally-export-parser';
import { parseBankStatement } from './bank-statement-parser';

const fixturesDir = path.join(__dirname, 'fixtures');

function loadSyntheticFixture(name: string): { documentId: string; rows: unknown[] } {
  const raw = fs.readFileSync(path.join(fixturesDir, name), 'utf8');
  const parsed = JSON.parse(raw) as { documentId: string; rows: unknown[] };
  expect(parsed.documentId).toBeTruthy();
  expect(Array.isArray(parsed.rows)).toBe(true);
  expect(parsed.rows.length).toBeGreaterThan(0);
  return parsed;
}

describe('ingestion parsers (synthetic fixtures)', () => {
  it('loads synthetic GSTR-2B fixture and stub throws Not implemented', () => {
    const fixture = loadSyntheticFixture('synthetic-gstr2b.json');
    expect(fixture.rows[0]).toEqual(
      expect.objectContaining({
        gstin: expect.any(String),
        invoiceNumber: expect.any(String),
        taxableValue: expect.any(Number),
      }),
    );
    expect(() => parseGstr2b(JSON.stringify(fixture))).toThrow(/Not implemented/i);
  });

  it('loads synthetic Form 26AS fixture and stub throws Not implemented', () => {
    const fixture = loadSyntheticFixture('synthetic-form26as.json');
    expect(fixture.rows[0]).toEqual(
      expect.objectContaining({
        pan: expect.any(String),
        tdsDeducted: expect.any(Number),
      }),
    );
    expect(() => parseForm26as(JSON.stringify(fixture))).toThrow(/Not implemented/i);
  });

  it('loads synthetic Tally fixture and stub throws Not implemented', () => {
    const fixture = loadSyntheticFixture('synthetic-tally.json');
    expect(fixture.rows[0]).toEqual(
      expect.objectContaining({
        voucherNumber: expect.any(String),
        amount: expect.any(Number),
      }),
    );
    expect(() => parseTallyExport(JSON.stringify(fixture))).toThrow(/Not implemented/i);
  });

  it('loads synthetic bank statement fixture and stub throws Not implemented', () => {
    const fixture = loadSyntheticFixture('synthetic-bank-statement.json');
    expect(fixture.rows[0]).toEqual(
      expect.objectContaining({
        date: expect.any(String),
        type: expect.stringMatching(/credit|debit/),
        amount: expect.any(Number),
      }),
    );
    expect(() => parseBankStatement(JSON.stringify(fixture))).toThrow(/Not implemented/i);
  });
});
