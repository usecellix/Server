import {
  extractGstins,
  extractPans,
  gstinCheckChar,
  GST_STATE_CODES,
  normalizeInvoiceKey,
  validateGstin,
  validateHsnSac,
  validateIfsc,
  validatePan,
  validateTan,
  validateUdyam,
} from './india-validators.util';

/** Build a GSTIN with a CORRECT check char from state + PAN (+ entity digit). */
function makeGstin(state: string, pan: string, entity = '1'): string {
  const base = `${state}${pan}${entity}Z`;
  return base + gstinCheckChar(base);
}

const PANS = [
  ['27', 'AABCA1234F'], ['29', 'AAACM5678K'], ['07', 'BBBPS4321Q'], ['33', 'CCCHD9876L'],
  ['24', 'DDDFR1111M'], ['19', 'EEETP2222N'], ['36', 'FFFCT3333P'], ['09', 'GGGAS4444R'],
  ['06', 'HHHPT5555S'], ['03', 'IIIBM6666T'], ['10', 'JJJAB7777U'], ['32', 'KKKGL8888V'],
] as const;
const VALID = PANS.map(([s, p]) => makeGstin(s, p));

function mutate(g: string, pos: number): string {
  const c = g[pos];
  const repl = /[0-9]/.test(c) ? (c === '7' ? '8' : '7') : c === 'A' ? 'B' : 'A';
  return g.slice(0, pos) + repl + g.slice(pos + 1);
}

describe('B10 india-validators', () => {
  describe('GSTIN', () => {
    it.each(VALID)('accepts known-good GSTIN %s', (g) => {
      const r = validateGstin(g);
      expect(r.status).toBe('valid');
      expect(r.stateName).toBeDefined();
    });

    it('fixtures: the sample GSTINs used in CELLIX_TEST_DATA.xlsx are valid', () => {
      for (const g of [
        '27AABCA1234F1Z9', '29AAACM5678K1ZV', '07BBBPS4321Q1ZA', '33CCCHD9876L1ZN',
        '24DDDFR1111M1ZA', '19EEETP2222N1ZY', '36FFFCT3333P1ZJ', '09GGGAS4444R1Z7',
      ]) {
        expect(validateGstin(g).status).toBe('valid');
      }
    });

    it('fails the checksum when a single character is changed (all 15 positions)', () => {
      // Positions 0-1 can turn into an invalid state, 2-11 into a bad PAN; only
      // assert "never valid" for every mutation, and "checksum_fail" where the
      // earlier checks still pass.
      let checksumFails = 0;
      for (const g of VALID) {
        for (let pos = 0; pos < 15; pos += 1) {
          const bad = mutate(g, pos);
          const r = validateGstin(bad);
          expect(r.status).not.toBe('valid');
          if (r.status === 'checksum_fail') checksumFails += 1;
        }
      }
      expect(checksumFails).toBeGreaterThan(60);
    });

    it('reports the expected check char for a typo in the last position', () => {
      const good = VALID[0];
      const bad = mutate(good, 14);
      const r = validateGstin(bad);
      expect(r.status).toBe('checksum_fail');
      expect(r.expectedCheckChar).toBe(good[14]);
    });

    it('classifies each failure kind', () => {
      expect(validateGstin('27AABCA1234F1Z').status).toBe('invalid_format'); // too short
      expect(validateGstin('NOT-A-GSTIN').status).toBe('invalid_format');
      expect(validateGstin('').status).toBe('invalid_format');
      expect(validateGstin(makeGstin('50', 'HHHPT5555S')).status).toBe('invalid_state');
      expect(validateGstin(makeGstin('27', 'AAAXA1234F')).status).toBe('invalid_pan'); // 4th char X
    });

    it('is tolerant of case/space/NBSP but never rewrites the identifier into another value', () => {
      const g = VALID[0];
      expect(validateGstin(` ${g.toLowerCase().slice(0, 7)} ${g.slice(7)} `).status).toBe('valid');
      // never auto-corrects: a typo'd value stays reported as the typo'd value
      expect(validateGstin(mutate(g, 14)).gstin).toBe(mutate(g, 14));
    });

    it('accepts 97 and 99 state codes and rejects 00 / 39', () => {
      expect(GST_STATE_CODES['97']).toBeDefined();
      expect(GST_STATE_CODES['99']).toBeDefined();
      expect(validateGstin(makeGstin('97', 'AABCA1234F')).status).toBe('valid');
      expect(validateGstin(makeGstin('00', 'AABCA1234F')).status).toBe('invalid_state');
      expect(validateGstin(makeGstin('39', 'AABCA1234F')).status).toBe('invalid_state');
    });

    it('extracts GSTINs from free text', () => {
      const g = VALID[2];
      expect(extractGstins(`Supplier ${g} invoice 233; again ${g}`)).toEqual([g]);
      expect(extractGstins('no identifiers here')).toEqual([]);
    });
  });

  describe('PAN', () => {
    it.each([
      ['AAAPA1234A', 'Individual'], ['AAACA1234A', 'Company'], ['AAAHA1234A', 'HUF'],
      ['AAAFA1234A', 'Firm / LLP'], ['AAATA1234A', 'Trust'], ['AAAAA1234A', 'AOP'],
      ['AAABA1234A', 'BOI'], ['AAALA1234A', 'Local Authority'],
      ['AAAJA1234A', 'Artificial Juridical Person'], ['AAAGA1234A', 'Government'],
    ])('%s is %s', (pan, type) => {
      const r = validatePan(pan);
      expect(r.status).toBe('valid');
      expect(r.entityType).toBe(type);
    });

    it('rejects bad format and unknown entity letters; lowercase is normalised, not "fixed" silently', () => {
      expect(validatePan('AAAPA12345').status).toBe('invalid_format');
      expect(validatePan('AAAXA1234A').status).toBe('invalid_entity_type');
      expect(validatePan('aaapa1234a').status).toBe('valid'); // normalised for comparison only
    });

    it('extracts a PAN from a sentence', () => {
      expect(extractPans('Vendor ABC PAN AABCA1234F paid on 5th')).toEqual(['AABCA1234F']);
    });
  });

  describe('TAN / IFSC / Udyam / HSN-SAC', () => {
    it('TAN', () => {
      expect(validateTan('MUMA12345B')).toBe(true);
      expect(validateTan('MUM12345B')).toBe(false);
    });
    it('IFSC', () => {
      expect(validateIfsc('SBIN0001234')).toBe(true);
      expect(validateIfsc('SBIN1001234')).toBe(false);
    });
    it('Udyam', () => {
      expect(validateUdyam('UDYAM-MH-12-1234567')).toBe(true);
      expect(validateUdyam('UDYAM-MH-12-123')).toBe(false);
    });
    it('HSN / SAC', () => {
      expect(validateHsnSac('8471')).toEqual({ valid: true, kind: 'HSN', code: '8471' });
      expect(validateHsnSac('84713010')).toMatchObject({ valid: true, kind: 'HSN' });
      expect(validateHsnSac('998314')).toMatchObject({ valid: true, kind: 'SAC' });
      expect(validateHsnSac('847').valid).toBe(false);
      expect(validateHsnSac('84A1').valid).toBe(false);
    });
  });

  describe('normalizeInvoiceKey', () => {
    it.each([
      ['INV/24-25/0233', '233'],
      ['Inv#233', '233'],
      ['0233', '233'],
      ['bill no 233 dt 12.04.24', '233'],
      ['INV-24-25-0234', '234'],
      ['OS-900', 'OS900'],
      ['OS900', 'OS900'],
      ['DL/24-25/77', 'DL77'],
      ['', ''],
      ['000', '0'],
    ])('%j → %j', (raw, key) => {
      expect(normalizeInvoiceKey(raw)).toBe(key);
    });

    it('does not eat a real number that merely looks like a year pair', () => {
      expect(normalizeInvoiceKey('INV/12-34')).toBe('1234');
    });
  });
});
