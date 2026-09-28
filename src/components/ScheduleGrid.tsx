import { useEffect, useMemo, useRef, useState, type CSSProperties, type PointerEvent as RPointerEvent } from 'react';
import {
  fellesferie,
  holidaysBetween,
  yearOf,
  formatDay,
  formatMonth,
  isoWeek,
  parseISO,
  startOfWeek,
  toISO,
  todayDay,
  type Day,
} from '../engine/dates';
import type { Project, Schedule, ScheduledPhase, ScheduledTask, ScheduleWarning } from '../engine/types';
import type { Action } from '../state/store';
import { DateInput, NumberInput, TextInput } from './cells';

export const ZOOMS = { day: 22, week: 9, month: 3.2 } as const;
export type Zoom = keyof typeof ZOOMS;

interface Props {
  project: Project;
  schedule: Schedule;
  dispatch: (a: Action) => void;
  zoom: Zoom;
  showNotes: boolean;
  showFellesferie: boolean;
  highlight: string | null;
  /** Scroll the timeline so this day is visible; `n` makes repeated requests re-trigger. */
  focus: { day: number; n: number } | null;
}

type Drag = { id: string; mode: 'move' | 'resize'; x0: number; dx: number };

export function ScheduleGrid({ project, schedule, dispatch, zoom, showNotes, showFellesferie, highlight, focus }: Props) {
  const dw = ZOOMS[zoom];
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  const [drag, setDrag] = useState<Drag | null>(null);
  const dragRef = useRef<Drag | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);

  const origin = useMemo(() => {
    const s = schedule.start ?? todayDay();
    return startOfWeek(s - 7);
  }, [schedule.start]);
  const lastDay = (schedule.end ?? origin + 365) + 21;
  const totalDays = lastDay - origin + 1;
  const width = totalDays * dw;
  const today = todayDay();
  const seaTrial = schedule.tasks.get('trials.seaTrial')?.start ?? null;

  const scrollToDay = (day: Day, behavior: ScrollBehavior = 'smooth') => {
    const el = scrollRef.current;
    if (!el) return;
    const left = parseFloat(getComputedStyle(el).getPropertyValue('--left-width')) || 0;
    el.scrollTo({ left: Math.max(0, (day - origin) * dw - (el.clientWidth - left) * 0.25), behavior });
  };

  // Keep the requested day in view (also after zooming).
  useEffect(() => {
    if (focus) scrollToDay(focus.day);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [focus, dw]);

  const warningsByTask = useMemo(() => {
    const m = new Map<string, ScheduleWarning[]>();
    for (const w of schedule.warnings) m.set(w.taskId, [...(m.get(w.taskId) ?? []), w]);
    return m;
  }, [schedule.warnings]);

  const baseline = project.baseline?.dates ?? null;

  const timelineBg: CSSProperties = {
    width,
    backgroundImage: `linear-gradient(to right, transparent ${5 * dw}px, var(--weekend) ${5 * dw}px), linear-gradient(to right, var(--grid) 1px, transparent 1px)`,
    backgroundSize: `${7 * dw}px 100%, ${7 * dw}px 100%`,
  };

  // --- header: months + weeks --------------------------------------------------
  const months: { start: Day; days: number }[] = [];
  {
    let d = origin;
    while (d <= lastDay) {
      const dt = new Date(d * 86_400_000);
      const next = Math.round(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 1) / 86_400_000);
      months.push({ start: d, days: Math.min(next, lastDay + 1) - d });
      d = next;
    }
  }
  const useHolidays = project.calendar !== 'none';
  const holidays = useHolidays ? holidaysBetween(origin, lastDay) : [];
  const ferie: [Day, Day][] = [];
  if (showFellesferie) for (let y = yearOf(origin); y <= yearOf(lastDay); y++) ferie.push(fellesferie(y));

  const weeks: Day[] = [];
  for (let d = origin; d <= lastDay; d += 7) weeks.push(d);

  // --- drag to move / resize ---------------------------------------------------
  const onBarPointerDown = (e: RPointerEvent, t: ScheduledTask, mode: Drag['mode']) => {
    if (t.start == null || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    (e.target as Element).setPointerCapture(e.pointerId);
    const d = { id: t.id, mode, x0: e.clientX, dx: 0 };
    dragRef.current = d;
    setDrag(d);
  };
  const onPointerMove = (e: RPointerEvent) => {
    const d = dragRef.current;
    if (!d) return;
    const next = { ...d, dx: e.clientX - d.x0 };
    dragRef.current = next;
    setDrag(next);
  };
  const onPointerUp = () => {
    const d = dragRef.current;
    dragRef.current = null;
    setDrag(null);
    if (!d) return;
    const t = schedule.tasks.get(d.id);
    const delta = Math.round(d.dx / dw);
    if (!t || t.start == null || delta === 0) return;
    if (d.mode === 'move') {
      dispatch({ type: 'setOverride', taskId: t.id, patch: { pinStart: toISO(t.start + delta) } });
    } else {
      dispatch({ type: 'setOverride', taskId: t.id, patch: { duration: Math.max(1, t.days + delta) } });
    }
  };

  const toggle = (id: string) =>
    setCollapsed((s) => {
      const n = new Set(s);
      if (n.has(id)) n.delete(id);
      else n.add(id);
      return n;
    });

  const line = (day: Day | null, cls: string, label?: string) =>
    day != null && day >= origin && day <= lastDay ? (
      <div className={`vline ${cls}`} style={{ left: (day - origin) * dw }} title={label} />
    ) : null;

  const renderPhase = (p: ScheduledPhase) => {
    const isCollapsed = collapsed.has(p.id);
    return (
      <div className="row phase-row" key={p.id} data-row={p.id}>
        <div className="left">
          <div className="c-wbs">{p.wbs}</div>
          <div className="c-name">
            <button className="twisty" onClick={() => toggle(p.id)} aria-label={isCollapsed ? 'Expand' : 'Collapse'}>
              {isCollapsed ? '▸' : '▾'}
            </button>
            <span className="phase-name">{p.name}</span>
            {p.direction === 'backward' && (
              <span className="dir" title="Planned backwards from the sea trial date">
                ◀ from sea trial
              </span>
            )}
          </div>
          <div className="c-date">{formatDay(p.start)}</div>
          <div className="c-date">{formatDay(p.end)}</div>
          <div className="c-num">{p.days ?? '–'}</div>
          <div className="c-num" />
          <div className="c-num">{p.hours || ''}</div>
          <div className="c-act" />
        </div>
        <div className="timeline" style={timelineBg}>
          {line(today, 'today')}
          {line(seaTrial, 'sea')}
          {p.start != null && p.end != null && (
            <div
              className="bar phase-bar"
              style={{ left: (p.start - origin) * dw, width: Math.max((p.end - p.start + 1) * dw, 2) }}
              title={`${p.name}\n${formatDay(p.start)} – ${formatDay(p.end)} (${p.days} days)`}
            />
          )}
        </div>
      </div>
    );
  };

  const renderTask = (t: ScheduledTask) => {
    const ws = warningsByTask.get(t.id);
    const err = ws?.some((w) => w.severity === 'error');
    if (t.kind === 'note') {
      return (
        <div className="row note-row" key={t.id} data-row={t.id}>
          <div className="left">
            <div className="c-wbs" />
            <div className="c-name note">• {t.name}</div>
            <div className="c-date" />
            <div className="c-date" />
            <div className="c-num" />
            <div className="c-num" />
            <div className="c-num" />
            <div className="c-act" />
          </div>
          <div className="timeline" style={timelineBg} />
        </div>
      );
    }

    const isDragging = drag?.id === t.id;
    const dDays = isDragging ? Math.round(drag.dx / dw) : 0;
    const start = t.start != null ? t.start + (isDragging && drag.mode === 'move' ? dDays : 0) : null;
    const days = isDragging && drag.mode === 'resize' ? Math.max(1, t.days + dDays) : t.days;

    const b = baseline?.[t.id];
    const bStart = b ? parseISO(b.start) : null;
    const bEnd = b ? parseISO(b.end) : null;
    const delta = bStart != null && t.start != null ? t.start - bStart : null;
    const endDelta = bEnd != null && t.end != null ? t.end - bEnd : null;

    const cls = [
      'row',
      'task-row',
      t.emphasis && 'emph',
      t.disabled && 'disabled',
      t.added && 'added',
      err && 'has-error',
      !err && ws && 'has-warning',
      highlight === t.id && 'highlight',
    ]
      .filter(Boolean)
      .join(' ');

    const barCls = [
      'bar',
      t.kind === 'milestone' || t.days === 0 ? 'milestone' : '',
      t.emphasis ? 'key' : '',
      t.vacation ? 'vacation' : '',
      t.pinned ? 'pinned' : '',
      err ? 'error' : '',
      isDragging ? 'dragging' : '',
    ].join(' ');

    return (
      <div className={cls} key={t.id} data-row={t.id}>
        <div className="left">
          <div className="c-wbs">
            <button className="wbs-link" title="Show bar" onClick={() => t.start != null && scrollToDay(t.start)}>
              {t.wbs}
            </button>
          </div>
          <div className="c-name" title={t.name}>
            <TextInput
              className="name-input"
              value={t.name}
              aria-label="Task name"
              onCommit={(name) => dispatch({ type: 'setOverride', taskId: t.id, patch: { name } })}
            />
          </div>
          <div className="c-date">
            <DateInput
              className={t.pinned === 'start' ? 'pinned' : ''}
              value={t.start != null ? toISO(t.start) : null}
              disabled={t.start == null && t.anchor.kind === 'date'}
              aria-label="Start"
              title={t.pinned === 'start' ? 'Pinned start date' : 'Change to pin this start date'}
              onCommit={(iso) => iso && dispatch({ type: 'setOverride', taskId: t.id, patch: { pinStart: iso } })}
            />
          </div>
          <div className="c-date">
            <DateInput
              className={t.pinned === 'end' ? 'pinned' : ''}
              value={t.end != null ? toISO(t.end) : null}
              disabled={t.start == null && t.anchor.kind === 'date'}
              aria-label="End"
              title={t.pinned === 'end' ? 'Pinned end date' : 'Change to pin this end date'}
              onCommit={(iso) => iso && dispatch({ type: 'setOverride', taskId: t.id, patch: { pinEnd: iso } })}
            />
          </div>
          <div className="c-num">
            <NumberInput
              className={t.modified ? 'modified' : ''}
              value={t.disabled ? 0 : t.days}
              disabled={t.disabled}
              aria-label="Days"
              title="Calendar days"
              onCommit={(n) => dispatch({ type: 'setOverride', taskId: t.id, patch: { duration: n ?? undefined } })}
            />
          </div>
          <div
            className={`c-num muted ${t.holidays.length ? 'has-holiday' : ''}`}
            title={`Work days (Mon–Fri${useHolidays ? ' excl. Norwegian public holidays' : ''})${t.holidays.length ? `\nHolidays: ${t.holidays.join(', ')}` : ''}`}
          >
            {t.workDays ?? ''}
          </div>
          <div className="c-num">
            <NumberInput
              value={t.hours}
              aria-label="Hours"
              title="Estimated hours"
              onCommit={(n) => dispatch({ type: 'setOverride', taskId: t.id, patch: { hours: n ?? undefined } })}
            />
          </div>
          <div className="c-act">
            {t.pinned && (
              <button
                className="icon pin"
                title="Pinned – click to release and follow the plan again"
                onClick={() => dispatch({ type: 'setOverride', taskId: t.id, patch: { pinStart: undefined, pinEnd: undefined } })}
              >
                📌
              </button>
            )}
            <button
              className={`icon skip ${t.disabled ? 'on' : ''}`}
              title={t.disabled ? 'Skipped – click to include again' : 'Skip this task (counts as 0 days)'}
              onClick={() => dispatch({ type: 'setOverride', taskId: t.id, patch: { disabled: !t.disabled } })}
            >
              {t.disabled ? '⊘' : '○'}
            </button>
            {(t.modified || t.pinned || t.disabled) && (
              <button className="icon" title="Reset to template" onClick={() => dispatch({ type: 'resetTask', taskId: t.id })}>
                ↺
              </button>
            )}
            {delta != null && (delta !== 0 || endDelta !== 0) && (
              <span
                className={`delta ${(endDelta ?? delta) > 0 ? 'late' : 'early'}`}
                title={`Compared with baseline: start ${delta >= 0 ? '+' : ''}${delta} d, end ${endDelta! >= 0 ? '+' : ''}${endDelta} d`}
              >
                {(endDelta ?? delta) > 0 ? '+' : ''}
                {endDelta ?? delta}d
              </span>
            )}
          </div>
        </div>
        <div className="timeline" style={timelineBg} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp}>
          {line(today, 'today')}
          {line(seaTrial, 'sea')}
          {bStart != null && bEnd != null && (
            <div className="baseline-bar" style={{ left: (bStart - origin) * dw, width: Math.max((bEnd - bStart + 1) * dw, 2) }} />
          )}
          {start != null && !t.disabled && (
            <div
              className={barCls}
              style={
                days === 0
                  ? { left: (start - origin) * dw - 6 }
                  : { left: (start - origin) * dw, width: Math.max(days * dw, 3) }
              }
              title={`${t.name}\n${formatDay(start)} – ${formatDay(start + Math.max(days - 1, 0))} (${days} days)${
                ws ? `\n⚠ ${ws.map((w) => w.message).join('\n⚠ ')}` : ''
              }${t.holidays.length ? `\nHolidays: ${t.holidays.join(', ')}` : ''}\nDrag to move · drag right edge to change duration`}
              onPointerDown={(e) => onBarPointerDown(e, t, 'move')}
            >
              {days > 0 && <span className="resize" onPointerDown={(e) => onBarPointerDown(e, t, 'resize')} />}
              {isDragging && (
                <span className="drag-label">
                  {drag.mode === 'move' ? formatDay(start) : `${days} d`}
                </span>
              )}
            </div>
          )}
          {!t.disabled && start != null && days * dw > 0 && t.pinned && (
            <span className="pin-mark" style={{ left: (start - origin) * dw - 5 }}>
              📌
            </span>
          )}
        </div>
      </div>
    );
  };

  return (
    <div className="grid-scroll" ref={scrollRef} style={{ '--dw': `${dw}px` } as CSSProperties}>
      <div className="grid" style={{ minWidth: `calc(var(--left-width) + ${width}px)` }}>
        <div className="row head-row head-1">
          <div className="left">
            <div className="c-wbs">WBS</div>
            <div className="c-name">Task</div>
            <div className="c-date">Start</div>
            <div className="c-date">End</div>
            <div className="c-num">Days</div>
            <div className="c-num" title="Work days">WD</div>
            <div className="c-num">Hours</div>
            <div className="c-act" />
          </div>
          <div className="timeline head" style={{ width }}>
            {months.map((m) => (
              <div key={m.start} className="month" style={{ left: (m.start - origin) * dw, width: m.days * dw }}>
                {m.days * dw > 44 ? formatMonth(m.start) : ''}
              </div>
            ))}
          </div>
        </div>
        <div className="row head-row head-2">
          <div className="left" />
          <div className="timeline head" style={{ width }}>
            {weeks.map((w) => (
              <div key={w} className="week" style={{ left: (w - origin) * dw, width: 7 * dw }}>
                {7 * dw >= 20 ? (zoom === 'day' ? `W${isoWeek(w)} · ${formatDay(w).slice(0, 6)}` : isoWeek(w)) : ''}
              </div>
            ))}
            {ferie.map(([a, b]) => (
              <div key={`f${a}`} className="ferie-head" style={{ left: (a - origin) * dw, width: (b - a + 1) * dw }} title="Fellesferie (weeks 28–30)">
                {(b - a + 1) * dw > 60 ? 'Fellesferie' : ''}
              </div>
            ))}
            {holidays.map((h) => (
              <div key={h.day} className="holiday-head" style={{ left: (h.day - origin) * dw, width: Math.max(dw, 4) }} title={`${h.name} – ${formatDay(h.day)}`} />
            ))}
          </div>
        </div>
        <div className="cal-overlay" aria-hidden>
          {ferie.map(([a, b]) => (
            <div key={`f${a}`} className="ferie" style={{ left: (a - origin) * dw, width: (b - a + 1) * dw }} />
          ))}
          {holidays.map((h) => (
            <div key={h.day} className="holiday" style={{ left: (h.day - origin) * dw, width: Math.max(dw, 2) }} />
          ))}
        </div>
        {schedule.phases.map((p) => [
          renderPhase(p),
          ...(collapsed.has(p.id) ? [] : p.tasks.filter((t) => showNotes || t.kind !== 'note').map(renderTask)),
        ])}
      </div>
    </div>
  );
}
