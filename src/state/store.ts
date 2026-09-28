import { useEffect, useMemo, useReducer } from 'react';
import { EXCEL_DEFAULT_DATES, MAX_REVISIONS, MIN_REVISIONS } from '../engine/template';
import type { KeyDateKey, Project, TaskOverride } from '../engine/types';
import { toISO, type Day } from '../engine/dates';

const STORAGE_KEY = 'dp-fmea-planner.v1';
const HISTORY_LIMIT = 100;

export interface Workspace {
  projects: Project[];
  currentId: string;
}

interface State {
  ws: Workspace;
  past: Workspace[];
  future: Workspace[];
}

export type Action =
  | { type: 'setKeyDate'; key: KeyDateKey; value: string | null }
  | { type: 'setOverride'; taskId: string; patch: Partial<TaskOverride> }
  | { type: 'resetTask'; taskId: string }
  | { type: 'resetAll' }
  | { type: 'setRevisions'; value: number }
  | { type: 'setOwnerReview'; value: boolean }
  | { type: 'setMeta'; name?: string; vessel?: string }
  | { type: 'setBaseline'; dates: Record<string, { start: Day; end: Day }> | null }
  | { type: 'newProject' }
  | { type: 'duplicateProject' }
  | { type: 'deleteProject' }
  | { type: 'selectProject'; id: string }
  | { type: 'importProject'; project: Project }
  | { type: 'undo' }
  | { type: 'redo' };

const uid = () => Math.random().toString(36).slice(2, 10);

export function newProject(name = 'New DP FMEA project'): Project {
  return {
    id: uid(),
    name,
    vessel: '',
    keyDates: { ...EXCEL_DEFAULT_DATES },
    classRevisions: 2,
    ownerReview: true,
    overrides: {},
    baseline: null,
    updatedAt: new Date().toISOString(),
  };
}

/** Accepts anything that looks like a project (e.g. an imported JSON file) and fills in defaults. */
export function normaliseProject(raw: unknown): Project {
  const p = (raw ?? {}) as Partial<Project>;
  const d = newProject();
  return {
    ...d,
    ...p,
    id: typeof p.id === 'string' ? p.id : d.id,
    keyDates: { ...d.keyDates, ...(p.keyDates ?? {}) },
    overrides: p.overrides && typeof p.overrides === 'object' ? p.overrides : {},
    classRevisions: Math.min(MAX_REVISIONS, Math.max(MIN_REVISIONS, Number(p.classRevisions) || 2)),
    ownerReview: p.ownerReview ?? true,
  };
}

function load(): Workspace {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const ws = JSON.parse(raw) as Workspace;
      const projects = (ws.projects ?? []).map(normaliseProject);
      if (projects.length) {
        return { projects, currentId: projects.some((p) => p.id === ws.currentId) ? ws.currentId : projects[0]!.id };
      }
    }
  } catch {
    // ignore – start fresh
  }
  const p = newProject('DP FMEA project (Excel template)');
  return { projects: [p], currentId: p.id };
}

function save(ws: Workspace) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(ws));
  } catch {
    // storage full or blocked – the app still works for this session
  }
}

function updateCurrent(ws: Workspace, fn: (p: Project) => Project): Workspace {
  return {
    ...ws,
    projects: ws.projects.map((p) => (p.id === ws.currentId ? { ...fn(p), updatedAt: new Date().toISOString() } : p)),
  };
}

function cleanOverride(o: TaskOverride): TaskOverride | null {
  const c: TaskOverride = {};
  for (const [k, v] of Object.entries(o) as [keyof TaskOverride, unknown][]) {
    if (v === undefined || v === null || v === '' || v === false) continue;
    (c as Record<string, unknown>)[k] = v;
  }
  return Object.keys(c).length ? c : null;
}

function apply(ws: Workspace, a: Action): Workspace {
  switch (a.type) {
    case 'setKeyDate':
      return updateCurrent(ws, (p) => ({ ...p, keyDates: { ...p.keyDates, [a.key]: a.value || (a.key === 'vesselDelivery' ? null : p.keyDates[a.key]) } }));
    case 'setOverride':
      return updateCurrent(ws, (p) => {
        const merged = { ...p.overrides[a.taskId], ...a.patch };
        // A new pin replaces the other kind of pin.
        if (a.patch.pinStart) delete merged.pinEnd;
        if (a.patch.pinEnd) delete merged.pinStart;
        const c = cleanOverride(merged);
        const overrides = { ...p.overrides };
        if (c) overrides[a.taskId] = c;
        else delete overrides[a.taskId];
        return { ...p, overrides };
      });
    case 'resetTask':
      return updateCurrent(ws, (p) => {
        const overrides = { ...p.overrides };
        delete overrides[a.taskId];
        return { ...p, overrides };
      });
    case 'resetAll':
      return updateCurrent(ws, (p) => ({ ...p, overrides: {} }));
    case 'setRevisions':
      return updateCurrent(ws, (p) => ({
        ...p,
        classRevisions: Math.min(MAX_REVISIONS, Math.max(MIN_REVISIONS, a.value)),
      }));
    case 'setOwnerReview':
      return updateCurrent(ws, (p) => ({ ...p, ownerReview: a.value }));
    case 'setMeta':
      return updateCurrent(ws, (p) => ({ ...p, name: a.name ?? p.name, vessel: a.vessel ?? p.vessel }));
    case 'setBaseline':
      return updateCurrent(ws, (p) => ({
        ...p,
        baseline: a.dates
          ? {
              takenAt: new Date().toISOString(),
              dates: Object.fromEntries(
                Object.entries(a.dates).map(([id, d]) => [id, { start: toISO(d.start), end: toISO(d.end) }]),
              ),
            }
          : null,
      }));
    case 'newProject': {
      const p = newProject();
      return { projects: [...ws.projects, p], currentId: p.id };
    }
    case 'duplicateProject': {
      const cur = ws.projects.find((p) => p.id === ws.currentId)!;
      const p = { ...structuredClone(cur), id: uid(), name: `${cur.name} (copy)`, updatedAt: new Date().toISOString() };
      return { projects: [...ws.projects, p], currentId: p.id };
    }
    case 'deleteProject': {
      const projects = ws.projects.filter((p) => p.id !== ws.currentId);
      if (!projects.length) {
        const p = newProject();
        return { projects: [p], currentId: p.id };
      }
      return { projects, currentId: projects[0]!.id };
    }
    case 'selectProject':
      return { ...ws, currentId: a.id };
    case 'importProject': {
      const p = normaliseProject(a.project);
      const exists = ws.projects.some((x) => x.id === p.id);
      const project = exists ? { ...p, id: uid() } : p;
      return { projects: [...ws.projects, project], currentId: project.id };
    }
    case 'undo':
    case 'redo':
      return ws;
  }
}

function reducer(state: State, a: Action): State {
  if (a.type === 'undo') {
    const prev = state.past.at(-1);
    if (!prev) return state;
    return { ws: prev, past: state.past.slice(0, -1), future: [state.ws, ...state.future] };
  }
  if (a.type === 'redo') {
    const [next, ...rest] = state.future;
    if (!next) return state;
    return { ws: next, past: [...state.past, state.ws], future: rest };
  }
  const ws = apply(state.ws, a);
  if (ws === state.ws) return state;
  return { ws, past: [...state.past, state.ws].slice(-HISTORY_LIMIT), future: [] };
}

export function useWorkspace() {
  const [state, dispatch] = useReducer(reducer, undefined, () => ({ ws: load(), past: [], future: [] }));
  useEffect(() => save(state.ws), [state.ws]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA')) return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        dispatch({ type: e.shiftKey ? 'redo' : 'undo' });
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        dispatch({ type: 'redo' });
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  const project = useMemo(
    () => state.ws.projects.find((p) => p.id === state.ws.currentId) ?? state.ws.projects[0]!,
    [state.ws],
  );

  return {
    workspace: state.ws,
    project,
    dispatch,
    canUndo: state.past.length > 0,
    canRedo: state.future.length > 0,
  };
}
