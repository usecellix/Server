/** Parse Indian accounting strings like `8533.98 Dr` for numeric sort. */
export function parseSortableValue(value: unknown): number | string {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'boolean') return value ? 1 : 0;

  const raw = String(value).trim();
  if (!raw) return '';

  const cleaned = raw
    .replace(/,/g, '')
    .replace(/[\s₹$€£]+/g, ' ')
    .replace(/\s*(Dr|Cr|DR|CR)\.?\s*$/i, '')
    .trim();

  const parsedDate = Date.parse(cleaned);
  if (!Number.isNaN(parsedDate) && /[-/]/.test(cleaned)) {
    return parsedDate;
  }

  const numeric = Number(cleaned);
  if (Number.isFinite(numeric)) return numeric;
  return raw;
}

export function compareSortValues(a: unknown, b: unknown): number {
  if (a === b) return 0;
  if (a === null || a === undefined || a === '') return 1;
  if (b === null || b === undefined || b === '') return -1;

  const av = parseSortableValue(a);
  const bv = parseSortableValue(b);

  if (typeof av === 'number' && typeof bv === 'number') return av - bv;
  return String(av).localeCompare(String(bv), undefined, { numeric: true, sensitivity: 'base' });
}
