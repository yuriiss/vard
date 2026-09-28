// Calendar arithmetic on whole days. A `Day` is the number of days since
// 1970-01-01 (UTC), so there are no time-zone or daylight-saving surprises.

export type Day = number;
/** ISO calendar date, `YYYY-MM-DD`. */
export type ISODate = string;

const MS_PER_DAY = 86_400_000;

export function parseISO(s: ISODate): Day {
  const [y, m, d] = s.split('-').map(Number);
  return Math.round(Date.UTC(y!, m! - 1, d!) / MS_PER_DAY);
}

export function isValidISO(s: string | null | undefined): s is ISODate {
  if (!s || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  return toISO(parseISO(s)) === s;
}

export function toISO(d: Day): ISODate {
  return new Date(d * MS_PER_DAY).toISOString().slice(0, 10);
}

export function toDate(d: Day): Date {
  return new Date(d * MS_PER_DAY);
}

/** 0 = Sunday … 6 = Saturday (1970-01-01 was a Thursday). */
export function weekday(d: Day): number {
  return (((d + 4) % 7) + 7) % 7;
}

export function isWeekend(d: Day): boolean {
  const w = weekday(d);
  return w === 0 || w === 6;
}

/** Monday on or before `d`. */
export function startOfWeek(d: Day): Day {
  return d - ((weekday(d) + 6) % 7);
}

/** Same as Excel NETWORKDAYS(a, b) without holidays: Mon–Fri between a and b inclusive. */
export function workDays(a: Day, b: Day): number {
  if (b < a) return -workDays(b, a);
  let n = 0;
  const fullWeeks = Math.floor((b - a + 1) / 7);
  n += fullWeeks * 5;
  for (let d = a + fullWeeks * 7; d <= b; d++) if (!isWeekend(d)) n++;
  return n;
}

export function addMonths(d: Day, months: number): Day {
  const dt = toDate(d);
  const y = dt.getUTCFullYear();
  const m = dt.getUTCMonth() + months;
  const day = dt.getUTCDate();
  // Clamp to the last day of the target month (31 Jan + 1 month = 28/29 Feb).
  const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return Math.round(Date.UTC(y, m, Math.min(day, last)) / MS_PER_DAY);
}

/** ISO-8601 week number. */
export function isoWeek(d: Day): number {
  const thursday = d - ((weekday(d) + 6) % 7) + 3;
  const jan1 = Math.round(Date.UTC(toDate(thursday).getUTCFullYear(), 0, 1) / MS_PER_DAY);
  return Math.floor((thursday - jan1) / 7) + 1;
}

const fmtShort = new Intl.DateTimeFormat('en-GB', {
  day: '2-digit',
  month: 'short',
  year: 'numeric',
  timeZone: 'UTC',
});

export function formatDay(d: Day | null | undefined): string {
  return d == null ? '–' : fmtShort.format(toDate(d));
}

const fmtMonth = new Intl.DateTimeFormat('en-GB', { month: 'short', year: 'numeric', timeZone: 'UTC' });
export function formatMonth(d: Day): string {
  return fmtMonth.format(toDate(d));
}

export function todayDay(): Day {
  const now = new Date();
  return Math.round(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) / MS_PER_DAY);
}
