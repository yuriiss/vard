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

// --- Norwegian public holidays -------------------------------------------------
// Same days as norskkalender.no lists under "Helligdager": the fixed dates plus the
// Easter-based ones. Computed, so it works for any year without a network call.

export interface Holiday {
  day: Day;
  name: string;
}

/** Easter Sunday (Gregorian), anonymous algorithm (Meeus/Jones/Butcher). */
export function easterSunday(year: number): Day {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return Math.round(Date.UTC(year, month - 1, day) / MS_PER_DAY);
}

const holidayCache = new Map<number, Holiday[]>();

/** Norwegian public holidays ("røde dager") of a year, sorted by date. */
export function norwegianHolidays(year: number): Holiday[] {
  const cached = holidayCache.get(year);
  if (cached) return cached;
  const fixed = (m: number, d: number) => Math.round(Date.UTC(year, m - 1, d) / MS_PER_DAY);
  const easter = easterSunday(year);
  const list: Holiday[] = [
    { day: fixed(1, 1), name: '1. nyttårsdag' },
    { day: easter - 7, name: 'Palmesøndag' },
    { day: easter - 3, name: 'Skjærtorsdag' },
    { day: easter - 2, name: 'Langfredag' },
    { day: easter, name: '1. påskedag' },
    { day: easter + 1, name: '2. påskedag' },
    { day: fixed(5, 1), name: 'Offentlig høytidsdag' },
    { day: fixed(5, 17), name: 'Grunnlovsdag' },
    { day: easter + 39, name: 'Kristi himmelfartsdag' },
    { day: easter + 49, name: '1. pinsedag' },
    { day: easter + 50, name: '2. pinsedag' },
    { day: fixed(12, 25), name: '1. juledag' },
    { day: fixed(12, 26), name: '2. juledag' },
  ];
  // Two holidays can fall on the same day (e.g. 17 May = 2. pinsedag in 2027).
  const byDay = new Map<Day, string>();
  for (const h of list) byDay.set(h.day, byDay.has(h.day) ? `${byDay.get(h.day)} / ${h.name}` : h.name);
  const result = [...byDay].map(([day, name]) => ({ day, name })).sort((x, y) => x.day - y.day);
  holidayCache.set(year, result);
  return result;
}

export function yearOf(d: Day): number {
  return toDate(d).getUTCFullYear();
}

/** Name of the Norwegian public holiday on `d`, if any. */
export function holidayName(d: Day): string | undefined {
  return norwegianHolidays(yearOf(d)).find((h) => h.day === d)?.name;
}

/** Holidays between a and b inclusive. */
export function holidaysBetween(a: Day, b: Day): Holiday[] {
  const out: Holiday[] = [];
  for (let y = yearOf(a); y <= yearOf(b); y++) for (const h of norwegianHolidays(y)) if (h.day >= a && h.day <= b) out.push(h);
  return out;
}

/**
 * Industrial common holiday ("fellesferie"): ISO weeks 28–30 of each year, the usual
 * shipyard/industry summer shutdown. Returns [firstDay, lastDay] (Mon–Sun).
 */
export function fellesferie(year: number): [Day, Day] {
  // ISO week 1 contains 4 January.
  const jan4 = Math.round(Date.UTC(year, 0, 4) / MS_PER_DAY);
  const week1 = startOfWeek(jan4);
  return [week1 + 27 * 7, week1 + 30 * 7 - 1];
}

/**
 * Mon–Fri between a and b inclusive, like Excel NETWORKDAYS(a, b).
 * With `holidays`, Norwegian public holidays on weekdays are excluded too
 * (NETWORKDAYS(a, b, holidays)).
 */
export function workDays(a: Day, b: Day, holidays = false): number {
  if (b < a) return -workDays(b, a, holidays);
  let n = 0;
  const fullWeeks = Math.floor((b - a + 1) / 7);
  n += fullWeeks * 5;
  for (let d = a + fullWeeks * 7; d <= b; d++) if (!isWeekend(d)) n++;
  if (holidays) for (const h of holidaysBetween(a, b)) if (!isWeekend(h.day)) n--;
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
