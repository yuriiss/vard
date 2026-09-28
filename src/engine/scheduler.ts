import {
  addMonths,
  fellesferie,
  formatDay,
  holidayName,
  holidaysBetween,
  isWeekend,
  parseISO,
  isValidISO,
  workDays,
  yearOf,
  type Day,
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

/**
 * Calculate every date in the plan.
 *
 * Each task is placed by its anchor (the old Excel formula) unless the user pinned its
 * start or end date; then everything that hangs off it moves with it. Durations are
 * calendar days, exactly like the Excel DAYS column.
 */
export function schedule(cfg: ProjectConfig, phases: PhaseDef[] = buildTemplate(cfg)): Schedule {
  if ((cfg.vacations ?? 'auto') === 'always') return computeOnce(cfg, phases, new Set());

  // Vacation blocks only count when they land in the summer. Moving one changes where the
  // others land, so repeat until nothing changes (a handful of rounds at most).
  const vacationIds = phases.flatMap((p) => p.tasks.filter((t) => t.vacation).map((t) => t.id));
  let off = new Set<string>();
  let result = computeOnce(cfg, phases, off);
  const seen = new Set<string>([key(off)]);
  for (let round = 0; round < 8; round++) {
    const next = new Set<string>();
    for (const id of vacationIds) {
      const t = result.tasks.get(id);
      if (!t || t.start == null || cfg.overrides[id]?.disabled) continue;
      // Where would the full block sit? Forward blocks keep their start, backward ones their end.
      const full = cfg.overrides[id]?.duration ?? t.duration;
      const backward = t.anchor.kind === 'before';
      const a = backward && off.has(id) ? t.start - full : t.start;
      const b = a + full - 1;
      if (!overlapsSummer(a, b)) next.add(id);
    }
    const k = key(next);
    if (k === key(off) || seen.has(k)) break;
    seen.add(k);
    off = next;
    result = computeOnce(cfg, phases, off);
  }
  return result;
}

const key = (s: Set<string>) => [...s].sort().join('|');

function overlapsSummer(a: Day, b: Day): boolean {
  for (let y = yearOf(a); y <= yearOf(b); y++) {
    const [fs, fe] = fellesferie(y);
    if (a <= fe && b >= fs) return true;
  }
  return false;
}

function computeOnce(cfg: ProjectConfig, phases: PhaseDef[], autoOff: Set<string>): Schedule {
  const useHolidays = (cfg.calendar ?? 'NO') === 'NO';
  const workdayStarts = cfg.workdayStarts ?? true;
  const isWorking = (d: Day) => !isWeekend(d) && !(useHolidays && holidayName(d));
  const nextWorking = (d: Day) => {
    while (!isWorking(d)) d++;
    return d;
  };
  const prevWorking = (d: Day) => {
    while (!isWorking(d)) d--;
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
    if (o?.disabled || t.kind === 'note' || autoOff.has(t.id)) return 0;
    return Math.max(0, Math.round(o?.duration ?? t.duration));
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
      span = { start: s, endEx: s + dur };
    } else if (o?.pinEnd && isValidISO(o.pinEnd)) {
      const e = parseISO(o.pinEnd);
      span = dur === 0 ? { start: e, endEx: e } : { start: e + 1 - dur, endEx: e + 1 };
    } else {
      const a = t.anchor;
      switch (a.kind) {
        case 'date': {
          const s = keyDay(a);
          if (s != null) span = { start: s, endEx: s + dur };
          else if (a.fallback) span = place(a.fallback, dur);
          break;
        }
        default:
          span = place(a, dur);
          break;
      }
    }

    visiting.delete(id);
    spans.set(id, span);
    return span;
  };

  /** Position from a link to another task. Key dates and pins are never moved. */
  function place(a: Anchor, dur: number): Span | null {
    switch (a.kind) {
      case 'after': {
        const r = resolve(a.ref);
        if (!r) return null;
        let s = r.endEx + (a.lag ?? 0);
        if (workdayStarts && dur > 0) s = nextWorking(s);
        return { start: s, endEx: s + dur };
      }
      case 'with': {
        const r = resolve(a.ref);
        if (!r) return null;
        let s = r.start + (a.lag ?? 0);
        if (workdayStarts && dur > 0) s = nextWorking(s);
        return { start: s, endEx: s + dur };
      }
      case 'before': {
        const r = resolve(a.ref);
        if (!r) return null;
        let s = r.start - (a.lag ?? 0) - dur;
        // Backward planning: start earlier rather than on a day off (never later – that
        // would overlap the task it has to finish before).
        if (workdayStarts && dur > 0) s = prevWorking(s);
        return { start: s, endEx: s + dur };
      }
      case 'date': {
        const s = keyDay(a);
        return s == null ? (a.fallback ? place(a.fallback, dur) : null) : { start: s, endEx: s + dur };
      }
      case 'none':
        return null;
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
    const st: ScheduledTask[] = p.tasks.map((t) => {
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
        disabled: !!o.disabled || autoOff.has(t.id),
        autoSkipped: autoOff.has(t.id) && !o.disabled,
        done: !!o.done,
        comment: o.comment,
        modified:
          (o.duration != null && o.duration !== t.duration) || (!!o.name?.trim() && o.name !== t.name) ||
          (o.hours != null && o.hours !== t.hours),
      };
      tasks.set(t.id, s);
      return s;
    });
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
