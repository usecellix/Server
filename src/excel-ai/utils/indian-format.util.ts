const INDIAN_LOCALE = 'en-IN';

export function formatIndianCurrency(value: number): string {
  return `₹${formatIndianNumber(value)}`;
}

export function formatIndianNumber(value: number, decimals = 0): string {
  return new Intl.NumberFormat(INDIAN_LOCALE, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(value);
}

export function formatIndianPercent(value: number, decimals = 0): string {
  return `${formatIndianNumber(value * 100, decimals)}%`;
}

export const INDIAN_CURRENCY_FORMAT = '₹#,##,##0';
export const INDIAN_CURRENCY_FORMAT_DECIMALS = '₹#,##,##0.00';
export const INDIAN_NUMBER_FORMAT = '#,##,##0';
export const INDIAN_NUMBER_FORMAT_DECIMALS = '#,##,##0.00';
export const INDIAN_DATE_FORMAT = 'dd-mm-yyyy';

export function parseIndianNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value !== 'string') return null;

  const cleaned = value.replace(/[\u20B9\u0024\u20AC\u00A3,\s%]/g, '').trim();
  if (!cleaned) return null;

  const isWrappedNegative = /^\(.+\)$/.test(cleaned);
  const numericCandidate = isWrappedNegative ? `-${cleaned.slice(1, -1)}` : cleaned;
  const parsed = Number(numericCandidate);
  return Number.isFinite(parsed) ? parsed : null;
}
