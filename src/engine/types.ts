import type { Day, ISODate } from './dates';

/** The dates you type in; everything else is calculated from them. */
export interface KeyDates {
  /** Contract / project start. Drives the forward chain (preparation, RFQ, kick-off). */
  projectStart: ISODate;
  /** Start of the DP FMEA First Pass (set manually, as in the Excel template). */
  firstPassStart: ISODate;
  /** First day of sea trials. Drives the backward chain (rev.0 → class approvals → proving trials). */
  seaTrial: ISODate;
  /** Vessel delivery date. Optional; project closure is 3 months after it. */
  vesselDelivery: ISODate | null;
}

export type KeyDateKey = keyof KeyDates;

/**
 * How a task's position is derived. This is the formula that sat in the START/END
 * columns of the Excel sheet, made explicit.
 */
export type Anchor =
  /** Starts on a key date (optionally shifted by days or months). */
  | { kind: 'date'; key: KeyDateKey; offsetDays?: number; offsetMonths?: number }
  /** Finish-to-start: starts the day after `ref` ends (Excel: `=D(prev)+1`). */
  | { kind: 'after'; ref: string; lag?: number }
  /** Start-to-start: starts together with `ref`. */
  | { kind: 'with'; ref: string; lag?: number }
  /** Backward planning: finishes the day before `ref` starts (Excel: `=C(next)-1`). */
  | { kind: 'before'; ref: string; lag?: number }
  /** Checklist line without dates (e.g. "Organize meeting with stakeholders…"). */
  | { kind: 'none' };

export type TaskKind = 'task' | 'note' | 'milestone';

export interface TaskDef {
  id: string;
  phaseId: string;
  name: string;
  /** Calendar days, as in the Excel DAYS column. */
  duration: number;
  hours?: number;
  anchor: Anchor;
  kind: TaskKind;
  /** Key deliverable (bold in the Excel sheet). */
  emphasis?: boolean;
  /** Buffer / vacation block that can be switched off. */
  vacation?: boolean;
  /** Task was added in the 2025-08-28 template update (blue text in Excel). */
  added?: boolean;
  /**
   * Extra "must not start before X has finished" checks. X is a task id or a phase id.
   * Used to detect where the backward chain (from sea trial) collides with the
   * forward chain (from project start).
   */
  mustFollow?: string[];
}

export interface PhaseDef {
  id: string;
  name: string;
  tasks: TaskDef[];
  /** 'forward' phases are planned from project start, 'backward' from sea trial. */
  direction: 'forward' | 'backward';
}

/** What the user changed on a single task. Keyed by task id, so it survives re-generation. */
export interface TaskOverride {
  duration?: number;
  hours?: number;
  name?: string;
  /** Pin the start date; the task no longer follows its predecessor. */
  pinStart?: ISODate;
  /** Pin the end date (duration kept, start moves). */
  pinEnd?: ISODate;
  /** Skip the task (e.g. no summer vacation in this project). Counts as 0 days. */
  disabled?: boolean;
  /** Mark as done. Purely informational. */
  done?: boolean;
  comment?: string;
}

export interface ProjectConfig {
  keyDates: KeyDates;
  /** Number of "DP FMEA Class approval rev.N" cycles. Excel template = 2. */
  classRevisions: number;
  /** Include "DP FMEA Owner review rev.1 (where applicable)". */
  ownerReview: boolean;
  overrides: Record<string, TaskOverride>;
}

export interface Project extends ProjectConfig {
  id: string;
  name: string;
  vessel: string;
  updatedAt: string;
  /** Snapshot of task dates to compare against ("what moved?"). */
  baseline?: { takenAt: string; dates: Record<string, { start: ISODate; end: ISODate }> } | null;
}

export interface ScheduledTask extends TaskDef {
  wbs: string;
  /** First day, or null if it cannot be scheduled (e.g. missing key date). */
  start: Day | null;
  /** Last day (inclusive). Equals start for milestones. */
  end: Day | null;
  /** Effective duration after overrides (0 when disabled). */
  days: number;
  workDays: number | null;
  pinned: 'start' | 'end' | null;
  disabled: boolean;
  done: boolean;
  comment?: string;
  /** Name/duration differ from the template. */
  modified: boolean;
}

export interface ScheduledPhase {
  id: string;
  name: string;
  wbs: string;
  direction: 'forward' | 'backward';
  start: Day | null;
  end: Day | null;
  days: number | null;
  hours: number;
  tasks: ScheduledTask[];
}

export interface ScheduleWarning {
  taskId: string;
  severity: 'error' | 'warning';
  message: string;
}

export interface Schedule {
  phases: ScheduledPhase[];
  tasks: Map<string, ScheduledTask>;
  warnings: ScheduleWarning[];
  /**
   * Days of slack between the end of the forward chain (First Pass) and the start
   * of the backward chain (Full DP FMEA rev.0). Negative = the plan does not fit.
   */
  buffer: number | null;
  start: Day | null;
  end: Day | null;
}
