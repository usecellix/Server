/**
 * Local stand-in for the colleague's B5 rule-table registry (C1.1).
 *
 * Every statutory value CELLIX uses in Stage 4 lives here — never inline in a lane —
 * with a version, a source and a CA-verification flag (§1.6 / §1.7 [CA-GATE]).
 * When B5 ships, replace `getRuleTable` with the registry call; the shapes below are
 * the contract the lanes code against.
 *
 * ⚠ caVerified: false on every table — the values are seeded from public sources and
 *   MUST be signed off by a CA before release (plan §1.7).
 */
export interface RuleTable<T> {
  id: string;
  version: string;
  source: string;
  caVerified: boolean;
  data: T;
}

export interface CalendarRule {
  return: string;
  /** Applies when the entity has these registrations. */
  needs: Array<'gst-monthly' | 'gst-qrmp' | 'tds' | 'company' | 'audit' | 'msme' | 'any'>;
  frequency: 'monthly' | 'quarterly' | 'annual' | 'fixed';
  /** Day of the month AFTER the period (monthly/quarterly), or explicit dd-mm list (fixed/annual, FY-relative). */
  dueDay?: number;
  /** For fixed dates: [monthIndex 1-12, day, yearOffset from FY start year]. */
  fixed?: Array<[number, number, number, string]>;
  /** Monthly exceptions: period month → [month, day, yearOffset]. */
  exceptions?: Record<number, [number, number, number]>;
  section: string;
}

const COMPLIANCE_CALENDAR: RuleTable<CalendarRule[]> = {
  id: 'compliance-calendar',
  version: '2026-27.v1',
  source: 'CGST Act/Rules (GSTR-1 r.59, GSTR-3B r.61, GSTR-9 s.44), Income-tax Act 1961 / Income-tax Act 2025 (TDS r.30, advance tax s.211, tax audit s.44AB, ITR s.139), MCA MSME-1 — seeded 2026-11-07',
  caVerified: false,
  data: [
    { return: 'GSTR-1', needs: ['gst-monthly'], frequency: 'monthly', dueDay: 11, section: 'CGST r.59' },
    { return: 'IFF (optional)', needs: ['gst-qrmp'], frequency: 'monthly', dueDay: 13, section: 'CGST r.59(2)' },
    { return: 'GSTR-3B', needs: ['gst-monthly'], frequency: 'monthly', dueDay: 20, section: 'CGST r.61' },
    { return: 'GSTR-3B (QRMP)', needs: ['gst-qrmp'], frequency: 'quarterly', dueDay: 22, section: 'CGST r.61 (22nd/24th by state)' },
    {
      return: 'TDS payment',
      needs: ['tds'],
      frequency: 'monthly',
      dueDay: 7,
      exceptions: { 3: [4, 30, 1] },
      section: 'IT r.30 (March → 30 April)',
    },
    {
      return: 'Advance tax',
      needs: ['any'],
      frequency: 'fixed',
      fixed: [
        [6, 15, 0, 'Q1 (15%)'],
        [9, 15, 0, 'Q2 (45%)'],
        [12, 15, 0, 'Q3 (75%)'],
        [3, 15, 1, 'Q4 (100%)'],
      ],
      section: 'IT s.211',
    },
    { return: 'MSME-1', needs: ['company', 'msme'], frequency: 'fixed', fixed: [[4, 30, 0, 'Oct–Mar (prev)'], [10, 31, 0, 'Apr–Sep']], section: 'MSMED s.405 order' },
    { return: 'GSTR-9 / 9C', needs: ['gst-monthly', 'gst-qrmp'], frequency: 'fixed', fixed: [[12, 31, 1, 'Annual']], section: 'CGST s.44' },
    { return: 'Tax audit report', needs: ['audit', 'company'], frequency: 'fixed', fixed: [[9, 30, 1, 'Annual']], section: 'IT s.44AB' },
    { return: 'ITR (non-audit)', needs: ['any'], frequency: 'fixed', fixed: [[7, 31, 1, 'Annual']], section: 'IT s.139(1)' },
  ],
};

const VALIDATION_PRESETS: RuleTable<Record<string, Array<string | number>>> = {
  id: 'validation-presets',
  version: '2026-11.v1',
  source: 'CGST rate notifications (rate slabs incl. 0.1% / 0.25% / 3%), Income-tax TDS sections — seeded 2026-11-07',
  caVerified: false,
  data: {
    'gst-rates': [0, 0.1, 0.25, 3, 5, 12, 18, 28],
    'tds-sections': ['192', '193', '194', '194A', '194C', '194H', '194I(a)', '194I(b)', '194J', '194Q', '195'],
    'yes-no': ['Yes', 'No'],
  },
};

const GST_TIME_LIMITS: RuleTable<{ itcDeadlineMonth: number; itcDeadlineDay: number; nearDays: number }> = {
  id: 'gst-time-limits',
  version: '2026-11.v1',
  source: 'CGST s.16(4) as amended by Finance Act 2022: ITC for a FY may be claimed till 30 November of the following FY (or annual return date, if earlier)',
  caVerified: false,
  data: { itcDeadlineMonth: 11, itcDeadlineDay: 30, nearDays: 30 },
};

const DRC01C: RuleTable<{ amountRupees: number; percent: number }> = {
  id: 'drc-01c',
  version: '2026-11.v1',
  source: 'CGST r.88D / Form DRC-01C — ITC in GSTR-3B exceeding GSTR-2B by more than ₹25 lakh and 20%',
  caVerified: false,
  data: { amountRupees: 2_500_000, percent: 20 },
};

export interface BankPack {
  bank: string;
  /** Narration prefix → mode. */
  modes: Array<[RegExp, string]>;
}

const BANK_PACKS: RuleTable<BankPack[]> = {
  id: 'bank-narration-packs',
  version: '2026-11.v1',
  source: 'Observed statement formats (SBI, HDFC, ICICI, Axis, Kotak) — internal fixtures',
  caVerified: true,
  data: [
    { bank: 'SBI', modes: [[/^BY TRANSFER-NEFT/i, 'NEFT'], [/^TO TRANSFER-UPI/i, 'UPI'], [/^BY CLEARING/i, 'CHQ']] },
    { bank: 'HDFC', modes: [[/^NEFT CR-/i, 'NEFT'], [/^UPI-/i, 'UPI'], [/^CHQ DEP/i, 'CHQ']] },
    { bank: 'ICICI', modes: [[/^NEFT-/i, 'NEFT'], [/^UPI\//i, 'UPI'], [/^CLG\//i, 'CHQ']] },
    { bank: 'Axis', modes: [[/^NEFT\//i, 'NEFT'], [/^UPI\/P2A/i, 'UPI'], [/^IMPS\/P2A/i, 'IMPS']] },
    { bank: 'Kotak', modes: [[/^NEFT /i, 'NEFT'], [/^UPI\//i, 'UPI'], [/^IMPS-/i, 'IMPS']] },
  ],
};

/** One TDS section: rates in %, thresholds in ₹ (single payment / aggregate in the FY). */
export interface TdsSectionRule {
  section: string;
  /** Section in the Income-tax Act, 2025 (effective 1 Apr 2026) — shown alongside (dual label). */
  newActRef: string;
  nature: string;
  /** Rate when the deductee is an individual / HUF (PAN 4th char P or H). */
  rateIndividual: number;
  /** Rate for every other deductee (company, firm, AOP…). */
  rateOthers: number;
  singleThreshold?: number;
  aggregateThreshold?: number;
}

const TDS_SECTIONS: RuleTable<{ noPanRate: number; sections: TdsSectionRule[] }> = {
  id: 'tds-sections',
  version: 'FY2026-27.v1',
  source:
    'Income-tax Act 1961 ss.194–194Q & 206AA (thresholds as amended by Finance Act 2025), with Income-tax Act 2025 s.393 table references — seeded 2026-11-07',
  caVerified: false,
  data: {
    noPanRate: 20,
    sections: [
      { section: '194C', newActRef: '393(1) Table Sl.6(i)', nature: 'Contractor / sub-contractor', rateIndividual: 1, rateOthers: 2, singleThreshold: 30_000, aggregateThreshold: 1_00_000 },
      { section: '194J', newActRef: '393(1) Table Sl.6(iii)', nature: 'Professional fees', rateIndividual: 10, rateOthers: 10, aggregateThreshold: 50_000 },
      { section: '194J(a)', newActRef: '393(1) Table Sl.6(iii)', nature: 'Technical services / royalty (film) / call centre', rateIndividual: 2, rateOthers: 2, aggregateThreshold: 50_000 },
      { section: '194H', newActRef: '393(1) Table Sl.1(ii)', nature: 'Commission / brokerage', rateIndividual: 2, rateOthers: 2, aggregateThreshold: 20_000 },
      { section: '194I(a)', newActRef: '393(1) Table Sl.2(ii)', nature: 'Rent — plant & machinery', rateIndividual: 2, rateOthers: 2, singleThreshold: 50_000 },
      { section: '194I(b)', newActRef: '393(1) Table Sl.2(ii)', nature: 'Rent — land, building, furniture', rateIndividual: 10, rateOthers: 10, singleThreshold: 50_000 },
      { section: '194A', newActRef: '393(1) Table Sl.5(iii)', nature: 'Interest (other than securities), non-bank payer', rateIndividual: 10, rateOthers: 10, aggregateThreshold: 10_000 },
      { section: '194', newActRef: '393(1) Table Sl.7', nature: 'Dividend', rateIndividual: 10, rateOthers: 10, aggregateThreshold: 10_000 },
      { section: '194Q', newActRef: '393(1) Table Sl.8(ii)', nature: 'Purchase of goods (buyer turnover > ₹10 cr)', rateIndividual: 0.1, rateOthers: 0.1, aggregateThreshold: 50_00_000 },
    ],
  },
};

const E_INVOICE: RuleTable<{ docNoPattern: string; docNoMaxLength: number; irnLength: number; ackNoLength: number }> = {
  id: 'e-invoice',
  version: '2026-11.v1',
  source:
    'GST e-invoice schema (INV-01, IRP spec): document no. max 16 chars, letters/digits plus / and -, not starting with 0, / or -; IRN = 64-char SHA-256 hex; Ack No. 15 digits',
  caVerified: false,
  data: { docNoPattern: '^[A-Za-z1-9][A-Za-z0-9/-]{0,15}$', docNoMaxLength: 16, irnLength: 64, ackNoLength: 15 },
};

const TABLES = {
  'tds-sections': TDS_SECTIONS,
  'e-invoice': E_INVOICE,
  'compliance-calendar': COMPLIANCE_CALENDAR,
  'validation-presets': VALIDATION_PRESETS,
  'gst-time-limits': GST_TIME_LIMITS,
  'drc-01c': DRC01C,
  'bank-narration-packs': BANK_PACKS,
} as const;

export type RuleTableId = keyof typeof TABLES;

export function getRuleTable<K extends RuleTableId>(id: K): (typeof TABLES)[K] {
  return TABLES[id];
}

/** "Rule: <id> v<version> · <source>" — shown next to every value a lane takes from a table. */
export function ruleCitation(id: RuleTableId): string {
  const t = TABLES[id];
  return `${t.id} v${t.version}${t.caVerified ? '' : ' (CA to verify)'} — ${t.source}`;
}
