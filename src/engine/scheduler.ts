import {
  addMonths,
  formatDay,
  holidayName,
  holidaysBetween,
  isWeekend,
  parseISO,
  isValidISO,
  workDays,
  yearOf,
  vacationWindows,
  type Day,
  type VacationWindow,
} from './dates';
import { buildTemplate } from './template';
import type {
  Anchor,
  PhaseDef,
  ProjectConfig,
  Schedule,
  ScheduledPhase,
  ScheduledTask,
  ScheduleWarning,
  TaskDef,
} from './types';

/** Internal span: `endEx` is the day *after* the last day, so a 0-day task has start === endEx. */
interface Span {
  start: Day;
  endEx: Day;
}

export class CycleError extends Error {}

export const DEFAULT_SUMMER = { start: '07-01', days: 28 } as const;
export const DEFAULT_WINTER = { start: '12-22', days: 14 } as const;

/**
 * Calculate every date in the plan.
 *
 * Each task is placed by its anchor (the old Excel formula) unless the user pinned its
 * start or end date; then everything that hangs off it moves with it. Durations are
 * calendar days, exactly like the Excel DAYS column.
 *
 * With `vacations: 'calendar'` company vacations are real calendar periods: work pauses
 * during them (a task that runs into the summer vacation is stretched by the overlap) and
 * they are shown as rows in the phases they interrupt, like the vacation rows in Excel.
 */
export function schedule(cfg: ProjectConfig, phases: PhaseDef[] = buildTemplate(cfg)): Schedule {
  const useHolidays = (cfg.calendar ?? 'NO') === 'NO';
  const workdayStarts = cfg.workdayStarts ?? true;
  const calendarVacations = (cfg.vacations ?? 'calendar') !== 'always';

  // Company vacation days, cached per year.
  const summer = cfg.summerVacation ?? DEFAULT_SUMMER;
  const winter = cfg.winterVacation ?? DEFAULT_WINTER;
  const vacCache = new Map<number, VacationWindow[]>();
  const windowsFor = (y: number) => {
    let w = vacCache.get(y);
    if (!w) {
      w = vacationWindows(y, y, summer, winter).filter((v) => yearOf(v.start) === y || yearOf(v.end) === y);
      vacCache.set(y, w);
    }
    return w;
  };
  const vacationAt = (d: Day): VacationWindow | undefined =>
    calendarVacations ? windowsFor(yearOf(d)).find((v) => d >= v.start && d <= v.end) : undefined;
  const inVacation = (d: Day) => vacationAt(d) !== undefined;

  const isWorking = (d: Day) => !isWeekend(d) && !(useHolidays && holidayName(d)) && !inVacation(d);
  const nextWorking = (d: Day) => {
    while (!isWorking(d)) d++;
    return d;
  };
  const prevWorking = (d: Day) => {
    while (!isWorking(d)) d--;
    return d;
  };
  const nextFree = (d: Day) => {
    while (inVacation(d)) d++;
    return d;
  };

  const defs = new Map<string, TaskDef>();
  for (const p of phases)
    for (const t of p.tasks) {
      if (defs.has(t.id)) throw new Error(`Duplicate task id "${t.id}"`);
      defs.set(t.id, t);
    }

  const spans = new Map<string, Span | null>();
  const visiting = new Set<string>();

  const keyDay = (a: Extract<Anchor, { kind: 'date' }>): Day | null => {
    const iso = cfg.keyDates[a.key] ?? null;
    if (!isValidISO(iso)) return null;
    let d = parseISO(iso);
    if (a.offsetMonths) d = addMonths(d, a.offsetMonths);
    return d + (a.offsetDays ?? 0);
  };

  const durationOf = (t: TaskDef): number => {
    const o = cfg.overrides[t.id];
    if (o?.disabled || t.kind === 'note' || (t.vacation && calendarVacations)) return 0;
    return Math.max(0, Math.round(o?.duration ?? t.duration));
  };

  /**
   * Does company vacation pause this task? Not for key-date tasks (sea trial, contract
   * period) and not for the "N weeks before sea trial" windows – those are fixed offsets.
   */
  const pauses = (t: TaskDef): boolean =>
    calendarVacations &&
    t.kind === 'task' &&
    !t.vacation &&
    t.anchor.kind !== 'date' &&
    !(t.anchor.kind === 'before' && t.anchor.ref === 'trials.seaTrial');

  /** Tasks up to this many days (meetings, uploads, short DOC tasks) are never split by a vacation. */
  const SHORT = 5;
  const firstVacationIn = (a: Day, bEx: Day): VacationWindow | undefined => {
    for (let d = a; d < bEx; d++) {
      const v = vacationAt(d);
      if (v) return v;
    }
    return undefined;
  };

  /** Forward: `dur` days of work from `start`; vacation days don't count. */
  const fitForward = (start: Day, dur: number, pause: boolean): Span => {
    if (!pause || dur === 0) return { start, endEx: start + dur };
    if (dur <= SHORT) {
      // Short task: move it after the vacation instead of splitting it.
      for (let v = firstVacationIn(start, start + dur); v; v = firstVacationIn(start, start + dur)) {
        start = workdayStarts ? nextWorking(v.end + 1) : v.end + 1;
      }
      return { start, endEx: start + dur };
    }
    let d = start;
    let left = dur;
    while (left > 0) {
      if (!inVacation(d)) left--;
      d++;
    }
    return { start, endEx: d };
  };

  /** Backward: `dur` days of work finishing before `endEx`; vacation days don't count. */
  const fitBackward = (endEx: Day, dur: number, pause: boolean): Span => {
    if (!pause || dur === 0) return { start: endEx - dur, endEx };
    if (dur <= SHORT) {
      for (let v = firstVacationIn(endEx - dur, endEx); v; v = firstVacationIn(endEx - dur, endEx)) endEx = v.start;
      return { start: endEx - dur, endEx };
    }
    let d = endEx - 1;
    let left = dur;
    while (left > 0) {
      if (!inVacation(d)) left--;
      d--;
    }
    return { start: d + 1, endEx };
  };

  const resolve = (id: string): Span | null => {
    if (spans.has(id)) return spans.get(id)!;
    const t = defs.get(id);
    if (!t) return null;
    if (visiting.has(id)) throw new CycleError(`Circular dependency at "${t.name}"`);
    visiting.add(id);

    const dur = durationOf(t);
    const o = cfg.overrides[id];
    let span: Span | null = null;

    if (t.kind === 'note') {
      span = null;
    } else if (o?.pinStart && isValidISO(o.pinStart)) {
      const s = parseISO(o.pinStart);
      span = { start: s, endEx: fitForward(s, dur, pauses(t)).endEx };
    } else if (o?.pinEnd && isValidISO(o.pinEnd)) {
      const e = parseISO(o.pinEnd);
      span = dur === 0 ? { start: e, endEx: e } : { start: fitBackward(e + 1, dur, pauses(t)).start, endEx: e + 1 };
    } else {
      const a = t.anchor;
      if (a.kind === 'date') {
        const s = keyDay(a);
        if (s != null) span = { start: s, endEx: s + dur };
        else if (a.fallback) span = place(a.fallback, dur, pauses({ ...t, anchor: a.fallback }));
      } else {
        span = place(a, dur, pauses(t));
      }
    }

    visiting.delete(id);
    spans.set(id, span);
    return span;
  };

  /** Position from a link to another task. Key dates and pins are never moved. */
  function place(a: Anchor, dur: number, pause: boolean): Span | null {
    const fwd = (s: Day): Span => {
      if (dur > 0) {
        if (workdayStarts) s = nextWorking(s);
        else if (pause) s = nextFree(s);
      }
      return fitForward(s, dur, pause);
    };
    switch (a.kind) {
      case 'after': {
        const r = resolve(a.ref);
        return r ? fwd(r.endEx + (a.lag ?? 0)) : null;
      }
      case 'with': {
        const r = resolve(a.ref);
        return r ? fwd(r.start + (a.lag ?? 0)) : null;
      }
      case 'before': {
        const r = resolve(a.ref);
        if (!r) return null;
        let endEx = r.start - (a.lag ?? 0);
        // Don't finish inside a vacation: finish before it instead.
        if (pause) while (dur > 0 && inVacation(endEx - 1)) endEx--;
        let span = fitBackward(endEx, dur, pause);
        // Backward planning: start earlier rather than on a day off (never later – that
        // would overlap the task it has to finish before).
        if (workdayStarts && dur > 0) {
          const p = prevWorking(span.start);
          if (p !== span.start) span = fitForward(p, dur, pause);
        }
        return span;
      }
      case 'date': {
        const s = keyDay(a);
        return s == null ? (a.fallback ? place(a.fallback, dur, pause) : null) : { start: s, endEx: s + dur };
      }
      case 'none':
        return null;
    }
  }

  /**
   * The stretch of calendar a task "owns": from where its predecessor lets it start (or from
   * its own start) to where its successor needs it done (or its own end). A vacation in that
   * stretch either paused the task or pushed it, so it belongs in the schedule next to it.
   */
  const ownedRange = (t: ScheduledTask): [Day, Day] | null => {
    const followsKickoff = t.anchor.kind === 'date' && !!t.anchor.fallback && !cfg.keyDates[t.anchor.key];
    if (t.start == null || t.end == null || t.disabled || t.days === 0 || !(pauses(t) || followsKickoff)) return null;
    let lo = t.start;
    let hi = t.end;
    const a = t.anchor;
    if (a.kind === 'after' || a.kind === 'with') {
      const r = spans.get(a.ref);
      if (r) lo = Math.min(lo, a.kind === 'after' ? r.endEx : r.start);
    } else if (a.kind === 'before') {
      const r = spans.get(a.ref);
      if (r) hi = Math.max(hi, r.start - 1);
    } else if (a.kind === 'date' && a.fallback && a.fallback.kind === 'after') {
      const r = spans.get(a.fallback.ref);
      if (r) lo = Math.min(lo, r.endEx);
    }
    return [lo, hi];
  };

  /** Add a read-only row for each company vacation that pauses or pushes a task of the phase. */
  function insertVacationRows(phaseId: string, rows: ScheduledTask[]) {
    const ranges = rows.map(ownedRange).filter((r): r is [Day, Day] => r != null);
    if (!ranges.length) return;
    const from = Math.min(...ranges.map((r) => r[0]));
    const to = Math.max(...ranges.map((r) => r[1]));
    const seen = new Set<string>();
    for (const v of vacationWindows(yearOf(from), yearOf(to), summer, winter)) {
      if (!ranges.some(([lo, hi]) => lo <= v.end && hi >= v.start) || seen.has(v.name)) continue;
      seen.add(v.name);
      const row: ScheduledTask = {
        id: `vacation.${phaseId}.${v.kind}.${v.start}`,
        phaseId,
        name: `${v.name} (${formatDay(v.start)} – ${formatDay(v.end)})`,
        duration: v.end - v.start + 1,
        anchor: { kind: 'none' },
        kind: 'task',
        vacation: true,
        calendarBlock: true,
        wbs: '',
        start: v.start,
        end: v.end,
        days: v.end - v.start + 1,
        workDays: null,
        holidays: [],
        pinned: null,
        disabled: false,
        pauseDays: 0,
        done: false,
        modified: false,
      };
      // Place it after the last row that starts on or before the vacation.
      let at = rows.length;
      for (let i = 0; i < rows.length; i++) if (rows[i]!.start != null && rows[i]!.start! > v.start) { at = i; break; }
      rows.splice(at, 0, row);
      tasks.set(row.id, row);
    }
  }

  // --- build output ---------------------------------------------------------------
  const out: ScheduledPhase[] = [];
  const tasks = new Map<string, ScheduledTask>();
  let phaseNo = 0;
  for (const p of phases) {
    // Project start is WBS 0.x like in Excel; the rest are numbered 1..n.
    const wbsBase = p.id === 'start' ? '0' : String(++phaseNo);
    let taskNo = 0;
    const st: ScheduledTask[] = p.tasks.filter((t) => !(t.vacation && calendarVacations)).map((t) => {
      const span = resolve(t.id);
      const o = cfg.overrides[t.id] ?? {};
      const days = durationOf(t);
      const start = span?.start ?? null;
      const end = span ? (days === 0 ? span.start : span.endEx - 1) : null;
      const s: ScheduledTask = {
        ...t,
        name: o.name?.trim() ? o.name : t.name,
        hours: o.hours ?? t.hours,
        wbs: t.kind === 'note' ? '' : `${wbsBase}.${++taskNo}`,
        start,
        end,
        days,
        workDays: start != null && end != null && days > 0 ? workDays(start, end, useHolidays) : null,
        holidays:
          useHolidays && start != null && end != null
            ? holidaysBetween(start, end)
                .filter((h) => !isWeekend(h.day))
                .map((h) => `${h.name} (${formatDay(h.day)})`)
            : [],
        pinned: o.pinStart ? 'start' : o.pinEnd ? 'end' : null,
        disabled: !!o.disabled,
        pauseDays: span ? span.endEx - span.start - days : 0,
        done: !!o.done,
        comment: o.comment,
        modified:
          (o.duration != null && o.duration !== t.duration) || (!!o.name?.trim() && o.name !== t.name) ||
          (o.hours != null && o.hours !== t.hours),
      };
      tasks.set(t.id, s);
      return s;
    });
    if (calendarVacations) insertVacationRows(p.id, st);
    const dated = st.filter((t) => t.start != null && !t.disabled);
    const start = dated.length ? Math.min(...dated.map((t) => t.start!)) : null;
    const end = dated.length ? Math.max(...dated.map((t) => t.end!)) : null;
    out.push({
      id: p.id,
      name: p.name,
      wbs: wbsBase,
      direction: p.direction,
      start,
      end,
      days: start != null && end != null ? end - start + 1 : null,
      hours: st.reduce((sum, t) => sum + (t.disabled ? 0 : (t.hours ?? 0)), 0),
      tasks: st,
    });
  }

  const phaseById = new Map(out.map((p) => [p.id, p]));
  const endOf = (ref: string): { end: Day | null; name: string } | null => {
    const t = tasks.get(ref);
    if (t) return { end: t.disabled ? (t.start == null ? null : t.start - 1) : t.end, name: t.name };
    const p = phaseById.get(ref);
    if (p) return { end: p.end, name: p.name };
    return null;
  };

  // --- consistency checks ---------------------------------------------------------
  const warnings: ScheduleWarning[] = [];
  const plural = (n: number) => `${n} day${n === 1 ? '' : 's'}`;
  for (const t of tasks.values()) {
    if (t.kind === 'note' || t.start == null || t.end == null) continue;
    const lastDay = t.days === 0 ? t.start - 1 : t.end;
    const a = t.anchor;

    // Pinned tasks can break the link they replaced – tell the user.
    if (t.pinned && a.kind !== 'none' && a.kind !== 'date') {
      const r = tasks.get(a.ref);
      if (r?.start != null && r.end != null) {
        const rLast = r.days === 0 ? r.start - 1 : r.end;
        if (a.kind === 'after' && t.start <= rLast) {
          warnings.push({
            taskId: t.id,
            severity: 'error',
            message: `Starts ${plural(rLast - t.start + 1)} before "${r.name}" is finished.`,
          });
        } else if (a.kind === 'before' && lastDay >= r.start) {
          warnings.push({
            taskId: t.id,
            severity: 'error',
            message: `Ends ${plural(lastDay - r.start + 1)} after "${r.name}" should start.`,
          });
        } else if (a.kind === 'after' && t.start > rLast + 1) {
          warnings.push({
            taskId: t.id,
            severity: 'warning',
            message: `Gap of ${plural(t.start - rLast - 1)} after "${r.name}".`,
          });
        }
      }
    }

    for (const ref of t.mustFollow ?? []) {
      const r = endOf(ref);
      if (r?.end != null && t.start <= r.end) {
        warnings.push({
          taskId: t.id,
          severity: 'error',
          message: `Starts ${plural(r.end - t.start + 1)} before "${r.name}" is finished.`,
        });
      }
    }
  }

  // Meetings, deliverables, the sea trial and hand-pinned dates should not start on a
  // Norwegian public holiday.
  if (useHolidays) {
    for (const t of tasks.values()) {
      if (t.start == null || t.disabled || t.kind === 'note') continue;
      const relevant = t.pinned || t.emphasis || /meeting|sea trial|upload|submit/i.test(t.name);
      const h = relevant ? holidayName(t.start) : undefined;
      if (h) {
        warnings.push({ taskId: t.id, severity: 'warning', message: `Starts on a public holiday: ${h} (${formatDay(t.start)}).` });
      }
    }
  }

  // Tasks placed before contract start.
  const projectStart = isValidISO(cfg.keyDates.projectStart) ? parseISO(cfg.keyDates.projectStart) : null;
  if (projectStart != null) {
    for (const t of tasks.values()) {
      if (t.start != null && t.start < projectStart && !t.disabled) {
        warnings.push({
          taskId: t.id,
          severity: 'error',
          message: `Starts ${plural(projectStart - t.start)} before project start.`,
        });
      }
    }
  }

  // Slack between forward chain (First Pass) and backward chain (rev.0).
  const fp = phaseById.get('firstPass');
  const rev0 = phaseById.get('rev0');
  const buffer = fp?.end != null && rev0?.start != null ? rev0.start - fp.end - 1 : null;

  const allDated = [...tasks.values()].filter((t) => t.start != null && !t.disabled);
  return {
    phases: out,
    tasks,
    warnings,
    buffer,
    start: allDated.length ? Math.min(...allDated.map((t) => t.start!)) : null,
    end: allDated.length ? Math.max(...allDated.map((t) => t.end!)) : null,
  };
}
