/**
 * Chart color schemes supported by CREATE_CHART / UPDATE_CHART.
 * Includes common named colors the model emits for "use green/red/…" requests.
 */
export type ChartColorScheme =
  | 'default'
  | 'blue'
  | 'grey'
  | 'blueGrey'
  | 'green'
  | 'red'
  | 'orange'
  | 'purple'
  | 'yellow';

const CANONICAL: Record<string, ChartColorScheme> = {
  default: 'default',
  blue: 'blue',
  grey: 'grey',
  gray: 'grey',
  bluegrey: 'blueGrey',
  green: 'green',
  /** Typo seen in user prompts ("greeen") — preserve intent. */
  greeen: 'green',
  red: 'red',
  orange: 'orange',
  purple: 'purple',
  yellow: 'yellow',
};

/** Hex fills used by Office.js chart series when applying a named scheme. */
export const CHART_COLOR_SCHEME_HEX: Record<Exclude<ChartColorScheme, 'default'>, string> = {
  blue: '#4472C4',
  grey: '#7F7F7F',
  blueGrey: '#5B9BD5',
  green: '#70AD47',
  red: '#C00000',
  orange: '#ED7D31',
  purple: '#7030A0',
  yellow: '#FFC000',
};

/**
 * Normalize a raw colorScheme from LLM output.
 * Returns undefined only when empty/absent — never silently drop a recognized color.
 */
export function normalizeChartColorScheme(raw: unknown): ChartColorScheme | undefined {
  if (typeof raw !== 'string') return undefined;
  const key = raw.trim().toLowerCase().replace(/[\s_-]+/g, '');
  if (!key) return undefined;
  return CANONICAL[key];
}
