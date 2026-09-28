import { useEffect, useMemo, useRef, useState } from 'react';
import { formatDay, todayDay } from './engine/dates';
import { schedule as computeSchedule, CycleError } from './engine/scheduler';
import { MAX_REVISIONS, MIN_REVISIONS } from './engine/template';
import type { Schedule } from './engine/types';
import { useWorkspace } from './state/store';
import { exportExcel, exportJson, readJsonFile } from './state/files';
import { DateInput, TextInput } from './components/cells';
import { ScheduleGrid, type Zoom } from './components/ScheduleGrid';

export default function App() {
  const { workspace, project, dispatch, canUndo, canRedo } = useWorkspace();
  const [zoom, setZoom] = useState<Zoom>('week');
  const [showNotes, setShowNotes] = useState(true);
  const [showFellesferie, setShowFellesferie] = useState(true);
  const [highlight, setHighlight] = useState<string | null>(null);
  const [focus, setFocus] = useState<{ day: number; n: number } | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const goTo = (day: number | null | undefined) => day != null && setFocus((f) => ({ day, n: (f?.n ?? 0) + 1 }));

  const result = useMemo((): { s: Schedule } | { error: string } => {
    try {
      return { s: computeSchedule(project) };
    } catch (e) {
      return { error: e instanceof CycleError || e instanceof Error ? e.message : String(e) };
    }
  }, [project]);

  // On open / project switch, show today (or the project start if today is outside the plan).
  useEffect(() => {
    if ('error' in result) return;
    const { start, end } = result.s;
    const today = todayDay();
    goTo(start != null && end != null && today >= start && today <= end ? today : start);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [project.id]);

  if ('error' in result) {
    return (
      <div className="fatal">
        <p>Could not calculate the schedule: {result.error}</p>
        <button onClick={() => dispatch({ type: 'undo' })}>Undo last change</button>{' '}
        <button onClick={() => dispatch({ type: 'resetAll' })}>Reset all task edits</button>
      </div>
    );
  }
  const s = result.s;
  const t = (id: string) => s.tasks.get(id);
  const kd = project.keyDates;
  const errors = s.warnings.filter((w) => w.severity === 'error');
  const editedCount = Object.keys(project.overrides).length;
  const totalHours = s.phases.reduce((n, p) => n + p.hours, 0);

  const jumpTo = (id: string) => {
    setHighlight(id);
    goTo(s.tasks.get(id)?.start);
    document.querySelector(`[data-row="${CSS.escape(id)}"]`)?.scrollIntoView({ block: 'center' });
    setTimeout(() => setHighlight((h) => (h === id ? null : h)), 2500);
  };

  const takeBaseline = () => {
    const dates: Record<string, { start: number; end: number }> = {};
    for (const task of s.tasks.values()) if (task.start != null && task.end != null) dates[task.id] = { start: task.start, end: task.end };
    dispatch({ type: 'setBaseline', dates });
  };

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="wordmark">vard</span>
          <span className="app-title">DP FMEA Planner</span>
        </div>
        <div className="project-picker">
          <select
            value={workspace.currentId}
            onChange={(e) => dispatch({ type: 'selectProject', id: e.target.value })}
            aria-label="Project"
          >
            {workspace.projects.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}
                {p.vessel ? ` – ${p.vessel}` : ''}
              </option>
            ))}
          </select>
          <button onClick={() => dispatch({ type: 'newProject' })} title="New project from the template">New</button>
          <button onClick={() => dispatch({ type: 'duplicateProject' })} title="Copy this project (e.g. to try a scenario)">Duplicate</button>
          <button
            onClick={() => confirm(`Delete "${project.name}"? This cannot be undone after reload.`) && dispatch({ type: 'deleteProject' })}
            title="Delete this project"
          >
            Delete
          </button>
        </div>
        <div className="toolbar">
          <button disabled={!canUndo} onClick={() => dispatch({ type: 'undo' })} title="Undo (Ctrl+Z)">↶ Undo</button>
          <button disabled={!canRedo} onClick={() => dispatch({ type: 'redo' })} title="Redo (Ctrl+Shift+Z)">↷ Redo</button>
          <span className="sep" />
          <button onClick={() => exportExcel(project, s)} title="Download an .xlsx with dates and a weekly Gantt">Export Excel</button>
          <button onClick={() => exportJson(project)} title="Save the project as a file you can share or re-import">Save file</button>
          <button onClick={() => fileRef.current?.click()} title="Open a saved .dpfmea.json file">Open file</button>
          <button onClick={() => window.print()}>Print</button>
          <input
            ref={fileRef}
            type="file"
            accept=".json,application/json"
            hidden
            onChange={async (e) => {
              const f = e.target.files?.[0];
              e.target.value = '';
              if (!f) return;
              try {
                dispatch({ type: 'importProject', project: (await readJsonFile(f)) as never });
              } catch {
                alert('That file is not a DP FMEA Planner project.');
              }
            }}
          />
        </div>
      </header>

      <section className="setup">
        <div className="card meta">
          <h2>Project</h2>
          <label>
            Name
            <TextInput value={project.name} onCommit={(name) => dispatch({ type: 'setMeta', name })} aria-label="Project name" />
          </label>
          <label>
            Vessel / hull no.
            <TextInput value={project.vessel} placeholder="e.g. Vard 9 32" onCommit={(vessel) => dispatch({ type: 'setMeta', vessel })} aria-label="Vessel" />
          </label>
        </div>

        <div className="card keydates">
          <h2>Key dates</h2>
          <label className="primary">
            Sea trial start
            <DateInput value={kd.seaTrial} onCommit={(v) => dispatch({ type: 'setKeyDate', key: 'seaTrial', value: v })} aria-label="Sea trial" />
            <small>Everything from rev.0 to proving trials is planned back from this.</small>
          </label>
          <label>
            Project start
            <DateInput value={kd.projectStart} onCommit={(v) => dispatch({ type: 'setKeyDate', key: 'projectStart', value: v })} aria-label="Project start" />
          </label>
          <label>
            First Pass start
            <DateInput value={kd.firstPassStart} onCommit={(v) => dispatch({ type: 'setKeyDate', key: 'firstPassStart', value: v })} aria-label="First Pass start" />
          </label>
          <label>
            Vessel delivery <span className="opt">(optional)</span>
            <DateInput allowEmpty value={kd.vesselDelivery} onCommit={(v) => dispatch({ type: 'setKeyDate', key: 'vesselDelivery', value: v })} aria-label="Vessel delivery" />
          </label>
        </div>

        <div className="card scope">
          <h2>Scope</h2>
          <div className="revs">
            <span>Class approval revisions</span>
            <div className="stepper">
              <button
                disabled={project.classRevisions <= MIN_REVISIONS}
                onClick={() => dispatch({ type: 'setRevisions', value: project.classRevisions - 1 })}
                aria-label="Remove revision"
              >
                −
              </button>
              <strong>{project.classRevisions}</strong>
              <button
                disabled={project.classRevisions >= MAX_REVISIONS}
                onClick={() => dispatch({ type: 'setRevisions', value: project.classRevisions + 1 })}
                aria-label="Add revision"
              >
                +
              </button>
            </div>
          </div>
          <small>
            rev.1{project.classRevisions > 1 ? ` … rev.${project.classRevisions}` : ''} · the last one issues R.
            {project.classRevisions + 1} and ends with the final Class approval
          </small>
          <label className="check">
            <input type="checkbox" checked={project.ownerReview} onChange={(e) => dispatch({ type: 'setOwnerReview', value: e.target.checked })} />
            Owner review rev.1 (parallel with Class rev.1)
          </label>
          <label className="check">
            <input
              type="checkbox"
              checked={project.calendar !== 'none'}
              onChange={(e) => dispatch({ type: 'setCalendar', value: e.target.checked ? 'NO' : 'none' })}
            />
            Norwegian public holidays (work days &amp; warnings)
          </label>
          <label className="check">
            <input type="checkbox" checked={showFellesferie} onChange={(e) => setShowFellesferie(e.target.checked)} />
            Show fellesferie (weeks 28–30)
          </label>
          <label className="check">
            <input type="checkbox" checked={showNotes} onChange={(e) => setShowNotes(e.target.checked)} />
            Show checklist notes
          </label>
        </div>

        <div className="card kpis">
          <h2>Result</h2>
          <dl>
            <dt>Full DP FMEA rev.0 starts</dt>
            <dd>{formatDay(t('rev0.collect')?.start)}</dd>
            <dt>Final Class approval</dt>
            <dd>{formatDay(t('classFinal.approval')?.end)}</dd>
            <dt>DP FMEA delivery goal</dt>
            <dd>{formatDay(t('trials.delivery')?.start)}</dd>
            <dt>Lessons learned done</dt>
            <dd>{formatDay(t('lessons.lessons')?.end)}</dd>
            <dt>Project closure</dt>
            <dd>{kd.vesselDelivery ? formatDay(t('closure.closure')?.start) : 'set vessel delivery'}</dd>
            <dt title="Days between end of First Pass and start of full DP FMEA rev.0">Buffer First Pass → rev.0</dt>
            <dd className={s.buffer != null && s.buffer < 0 ? 'bad' : 'good'}>
              {s.buffer == null ? '–' : `${s.buffer} days (${(s.buffer / 7).toFixed(1)} w)`}
            </dd>
            <dt>Estimated hours</dt>
            <dd>{totalHours}</dd>
          </dl>
        </div>
      </section>

      {s.warnings.length > 0 && (
        <section className={`warnings ${errors.length ? 'err' : ''}`}>
          <strong>
            {errors.length ? `${errors.length} conflict${errors.length > 1 ? 's' : ''}` : ''}
            {errors.length && s.warnings.length > errors.length ? ', ' : ''}
            {s.warnings.length > errors.length ? `${s.warnings.length - errors.length} note(s)` : ''}
          </strong>
          <ul>
            {s.warnings.slice(0, 8).map((w, i) => (
              <li key={i} className={w.severity}>
                <button className="link" onClick={() => jumpTo(w.taskId)}>
                  {s.tasks.get(w.taskId)?.wbs} {s.tasks.get(w.taskId)?.name}
                </button>{' '}
                – {w.message}
              </li>
            ))}
            {s.warnings.length > 8 && <li>… and {s.warnings.length - 8} more</li>}
          </ul>
        </section>
      )}

      <section className="gridbar">
        <div className="legend">
          <span><i className="lg phase" /> Phase</span>
          <span><i className="lg task" /> Task</span>
          <span><i className="lg key" /> Deliverable</span>
          <span><i className="lg vacation" /> Vacation / buffer</span>
          <span><i className="lg sea" /> Sea trial</span>
          <span><i className="lg today" /> Today</span>
          {project.calendar !== 'none' && <span><i className="lg holiday" /> Public holiday</span>}
          {showFellesferie && <span><i className="lg ferie" /> Fellesferie</span>}
          {project.baseline && <span><i className="lg baseline" /> Baseline</span>}
        </div>
        <div className="controls">
          <span className="hint">Edit any date or duration – or drag a bar. 📌 = pinned date.</span>
          {editedCount > 0 && (
            <button onClick={() => confirm('Remove all task edits (pins, durations, skips)?') && dispatch({ type: 'resetAll' })}>
              Reset {editedCount} edit{editedCount > 1 ? 's' : ''}
            </button>
          )}
          {project.baseline ? (
            <>
              <span className="muted">Baseline {project.baseline.takenAt.slice(0, 10)}</span>
              <button onClick={takeBaseline}>Update baseline</button>
              <button onClick={() => dispatch({ type: 'setBaseline', dates: null })}>Clear</button>
            </>
          ) : (
            <button onClick={takeBaseline} title="Freeze the current dates to see what moves later">Set baseline</button>
          )}
          <div className="seg" role="group" aria-label="Go to">
            <button onClick={() => goTo(s.start)}>Start</button>
            <button onClick={() => goTo(todayDay())}>Today</button>
            <button onClick={() => goTo(t('rev0.collect')?.start)}>rev.0</button>
            <button onClick={() => goTo(t('trials.seaTrial')?.start)}>Sea trial</button>
          </div>
          <div className="seg" role="group" aria-label="Zoom">
            {(['day', 'week', 'month'] as Zoom[]).map((z) => (
              <button key={z} className={zoom === z ? 'on' : ''} onClick={() => setZoom(z)}>
                {z[0]!.toUpperCase() + z.slice(1)}
              </button>
            ))}
          </div>
        </div>
      </section>

      <ScheduleGrid project={project} schedule={s} dispatch={dispatch} zoom={zoom} showNotes={showNotes} showFellesferie={showFellesferie} highlight={highlight} focus={focus} />

      <footer className="foot">
        <span className="legal">Vard Electro AS</span> · DP FMEA delivery planning · data is stored in this browser – use “Save file” to share
      </footer>
    </div>
  );
}
