import {
  bandOf,
  invoiceSimilarity,
  matchOneToOne,
  normalizePartyName,
  similarity,
  uniqueSubsetSum,
  withinTolerance,
} from './matching-engine';

describe('P2.1 matching engine', () => {
  it.each([
    ['M/s ABC Private Limited.', 'abc'],
    ['ABC Pvt Ltd', 'abc'],
    ['Kumar & Sons', 'kumar and sons'],
    ['XYZ Traders LLP', 'xyz traders'],
  ])('normalises %j → %j', (a, b) => expect(normalizePartyName(a)).toBe(b));

  it('scores the T_FuzzyA / T_FuzzyB pairs', () => {
    expect(similarity('ABC Pvt Ltd', 'ABC Private Limited')).toBeGreaterThanOrEqual(95);
    expect(similarity('M/s Sharma Enterprises', 'Sharma Enterprises')).toBe(100);
    expect(similarity('Kumar & Sons', 'Kumar and Sons')).toBe(100);
    expect(similarity('Delta Logistics Limited', 'Delta Logistic Ltd.')).toBeGreaterThanOrEqual(80);
    expect(similarity('Omega Steel Industries', 'Omega Steels')).toBeGreaterThanOrEqual(80);
    expect(similarity('Zenith Traders', 'Quick Mart')).toBeLessThan(80);
  });

  it('bands are configurable', () => {
    expect(bandOf(96)).toBe('high');
    expect(bandOf(85)).toBe('review');
    expect(bandOf(50)).toBe('none');
    expect(bandOf(85, { high: 80, review: 60 })).toBe('high');
  });

  it('one-to-one: each right row used once, unmatched left gets none, deterministic', () => {
    const left = ['ABC Pvt Ltd', 'ABC Private Limited', 'Zenith Traders'];
    const right = ['ABC Pvt. Ltd.'];
    const run = () => matchOneToOne(left, right, (l, r) => similarity(l, r));
    const out = run();
    expect(out.filter((c) => c.right !== null)).toHaveLength(1);
    expect(out[0].rightIndex).toBe(0);
    expect(out[2].band).toBe('none');
    expect(run()).toEqual(out);
  });

  it('tolerance: absolute and percent', () => {
    expect(withinTolerance(1000, 1050, { absolute: 100 })).toBe(true);
    expect(withinTolerance(1000, 1050, { absolute: 1 })).toBe(false);
    expect(withinTolerance(1000, 1009, { percent: 1 })).toBe(true);
    expect(withinTolerance(1000, 1000.001)).toBe(true);
  });

  it('many-to-one: unique subset found, ambiguity refused', () => {
    expect(uniqueSubsetSum([1000, 2000, 3000, 750], 6000)).toEqual([0, 1, 2]);
    expect(uniqueSubsetSum([1000, 2000, 1000, 2000], 3000)).toBeNull();
    expect(uniqueSubsetSum([10, 20], 99)).toBeNull();
  });

  it('invoice similarity uses the shared key', () => {
    expect(invoiceSimilarity('INV/24-25/0233', '233')).toBe(100);
    expect(invoiceSimilarity('OS-900', 'OS900')).toBe(100);
  });
});
