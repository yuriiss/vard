// The DP FMEA delivery process, transcribed from the Excel Gantt template
// ("DP FMEA process schedule", last update 2025-08-28).
//
// Two chains meet in the middle:
//   * forward  – from project start: preparation → RFQ → kick-off; First Pass from its own start date
//   * backward – from sea trial: proving trials ← class approval rev.N … rev.1 ← full DP FMEA rev.0
// Owner review runs in parallel with class approval rev.1.

import type { Anchor, PhaseDef, ProjectConfig, TaskDef, TaskKind } from './types';

interface Spec {
  key: string;
  name: string;
  days?: number;
  hours?: number;
  kind?: TaskKind;
  /** Explicit anchor; the task is then not part of the phase's sequential chain. */
  anchor?: Anchor;
  emphasis?: boolean;
  vacation?: boolean;
  added?: boolean;
  mustFollow?: string[];
  /** Override the id (used for tasks that move between phases, e.g. final class approval). */
  id?: string;
}

const MEETING_NOTE = 'Organize meeting with stakeholders to clarify, agree and update';

const note = (key: string, name = MEETING_NOTE): Spec => ({ key, name, kind: 'note' });

function makeTask(phaseId: string, s: Spec, anchor: Anchor): TaskDef {
  const kind = s.kind ?? 'task';
  return {
    id: s.id ?? `${phaseId}.${s.key}`,
    phaseId,
    name: s.name,
    duration: kind === 'note' ? 0 : (s.days ?? 1),
    hours: s.hours,
    anchor: kind === 'note' ? { kind: 'none' } : anchor,
    kind,
    emphasis: s.emphasis,
    vacation: s.vacation,
    added: s.added,
    mustFollow: s.mustFollow,
  };
}

/** Tasks follow each other; the first one hangs off `head`. */
function forward(id: string, name: string, head: Anchor, specs: Spec[]): PhaseDef {
  let prev: string | null = null;
  const tasks = specs.map((s) => {
    if (s.kind === 'note') return makeTask(id, s, { kind: 'none' });
    const anchor: Anchor = s.anchor ?? (prev ? { kind: 'after', ref: prev } : head);
    const t = makeTask(id, s, anchor);
    if (!s.anchor) prev = t.id;
    return t;
  });
  return { id, name, tasks, direction: 'forward' };
}

/** Tasks are planned back from `tailRef`: the last one finishes the day before `tailRef` starts. */
function backward(id: string, name: string, tailRef: string, specs: Spec[]): PhaseDef {
  let next = tailRef;
  const tasks: TaskDef[] = new Array(specs.length);
  for (let i = specs.length - 1; i >= 0; i--) {
    const s = specs[i]!;
    if (s.kind === 'note') {
      tasks[i] = makeTask(id, s, { kind: 'none' });
      continue;
    }
    const anchor: Anchor = s.anchor ?? { kind: 'before', ref: next };
    const t = makeTask(id, s, anchor);
    if (!s.anchor) next = t.id;
    tasks[i] = t;
  }
  return { id, name, tasks, direction: 'backward' };
}

export const UPLOAD = 'Ask VDE/VVT/VVT to upload';

/** Tasks of "DP FMEA Class approval rev.N". The last revision also carries the final approval. */
function classRevisionSpecs(n: number, isLast: boolean): Spec[] {
  const next = n + 1;
  const specs: Spec[] =
    n === 1
      ? [
          { key: 'upload', name: `${UPLOAD} for Class approval`, days: 1, hours: 1 },
          { key: 'approval', name: '3-8 weeks Class approval', days: 21, hours: 1 },
          { key: 'respond', name: '2 weeks for VARD to respond comments & update & QC', days: 14, hours: 31 },
          note('stakeholders'),
          { key: 'vacation', name: '4 weeks summer vacation', days: 28, vacation: true },
          {
            key: 'doc',
            name: 'DOC: DP FMEA - R.2 (1 week after all Class [&Owner] comments are resolved)',
            days: 3,
            hours: 1,
            emphasis: true,
            mustFollow: ['owner.respond'],
          },
          { key: 'qa', name: 'Perform QA on DP FMEA update from R.1 to R.2', days: 2, hours: 6 },
        ]
      : [
          { key: 'upload', name: `${UPLOAD} for Class re-approval`, days: 3, hours: 1 },
          ...(n === 2
            ? [
                {
                  key: 'testProcedure',
                  name: 'Submit test procedure to OEMs and Suppliers',
                  days: 1,
                  hours: 4,
                  anchor: { kind: 'after', ref: 'rev2.upload' },
                } satisfies Spec,
              ]
            : []),
          { key: 'approval', name: '1-3 weeks Class approval', days: 14, hours: 1 },
          { key: 'respond', name: '2 weeks for VARD to respond comments & update & QC', days: 14, hours: 29 },
          note('stakeholders'),
          {
            key: 'doc',
            name: `DOC: DP FMEA - R.${next} (1 week after all Class comments are resolved)`,
            days: 2,
            hours: 1,
            emphasis: true,
          },
          { key: 'qa', name: `Perform QA on DP FMEA update from R.${n} to R.${next}`, days: 3, hours: 6 },
        ];
  if (isLast) {
    // Stable ids so edits survive adding/removing revisions.
    specs.push(
      { key: 'finalUpload', id: 'classFinal.upload', name: `${UPLOAD} for Class re-approval`, days: 1, hours: 1 },
      { key: 'finalApproval', id: 'classFinal.approval', name: '1-3 weeks Class approval', days: 14, hours: 1 },
    );
  }
  return specs;
}

export const MIN_REVISIONS = 1;
export const MAX_REVISIONS = 9;

/** Build the full process for a given configuration. */
export function buildTemplate(cfg: Pick<ProjectConfig, 'classRevisions' | 'ownerReview'>): PhaseDef[] {
  const revs = Math.min(MAX_REVISIONS, Math.max(MIN_REVISIONS, Math.round(cfg.classRevisions)));
  const phases: PhaseDef[] = [];

  phases.push(
    forward('start', 'Project start', { kind: 'date', key: 'projectStart' }, [
      { key: 'period', name: 'Project start (period between Contract and Project Start)', days: 14 },
    ]),
  );

  phases.push(
    forward('prep', 'DP FMEA Preparation and planning', { kind: 'after', ref: 'start.period' }, [
      { key: 'exchange', name: 'Prepare DP FMEA Exchange folder in VPP', days: 1, hours: 2 },
      { key: 'plmStructure', name: 'Prepare Delivery Structure in PLM', days: 1, hours: 2 },
      { key: 'plmPlan', name: 'Prepare Delivery Plan in PLM', days: 1, hours: 2 },
      { key: 'docList', name: 'DOC: Prepare vessel specific list of documentation needed', days: 4, hours: 16 },
      { key: 'dependencies', name: 'Add known dependencies (Full DP FMEA) in delivery plan', days: 6, hours: 16 },
      { key: 'kickoff', name: 'Kick-off meeting between VE & VDE', days: 2, hours: 2 },
      note('n1', 'Discuss planning, time frame, duration, project particularities'),
      note('n2', 'Identify risk areas and systems subject to First Pass analysis'),
      note('n3', 'Review concept DP philosophy and Electrical philosophy'),
      note('n4', 'Review redundancy concept'),
      { key: 'mom', name: 'MOM: Kick-off meeting minutes VE & VDE', days: 1, hours: 2 },
    ]),
  );

  phases.push(
    forward('rfq', 'DP FMEA RFQ', { kind: 'after', ref: 'prep.mom' }, [
      { key: 'prepare', name: 'DOC: Prepare RFQ', days: 2, hours: 4 },
      { key: 'submit', name: 'Submit RFQ to Procurement', days: 1, hours: 1 },
      { key: 'wait', name: 'Wait 2 weeks (procurement standard process)', days: 14, hours: 1 },
      { key: 'evaluate', name: 'Evaluate offers', days: 5, hours: 14 },
      { key: 'nominate', name: 'Nominate DP FMEA supplier', days: 1, hours: 4 },
      { key: 'consult', name: 'Consult VP PM before official nomination', days: 1, hours: 1 },
      { key: 'tss', name: 'DOC: DP FMEA TSS', days: 3, hours: 2 },
      { key: 'po', name: 'Issue PO', days: 3, hours: 1 },
      { key: 'waitPo', name: 'Wait 1 week after PO is issued before next step', days: 7, hours: 1 },
    ]),
  );

  phases.push(
    forward('kickoff', 'DP FMEA Kick-off meetings', { kind: 'after', ref: 'rfq.waitPo' }, [
      { key: 'meeting', name: 'Kick-off meeting between VE & VDE & DP FMEA maker', days: 1, hours: 4 },
      { key: 'mom', name: 'MOM: Kick-off meeting minutes VE & VDE & DP FMEA maker', days: 3, hours: 2 },
      {
        key: 'matrix',
        name: 'DOC: Communication matrix. Prepare and distribute with project organization.',
        days: 3,
        hours: 2,
      },
      { key: 'mobilise', name: 'Wait approx. 1 week (to allocate enough mobilization time)', days: 9, hours: 1 },
      {
        key: 'classMeeting',
        name: 'Kick-off meeting between VE & VDE & DP FMEA maker & Class (Optional)',
        days: 2,
        hours: 4,
      },
      {
        key: 'classMom',
        name: 'MOM: Kick-off meeting minutes VE & VDE & DP FMEA maker & Class (Optional)',
        days: 1,
        hours: 2,
      },
      note('stakeholders'),
    ]),
  );

  phases.push(
    forward(
      'firstPass',
      'DP FMEA First Pass',
      { kind: 'date', key: 'firstPassStart', fallback: { kind: 'after', ref: 'kickoff.classMom' } },
      [
      {
        key: 'mark',
        name: 'DOC: Mark drawings that are subject for DP FMEA First Pass',
        days: 3,
        hours: 8,
        mustFollow: ['kickoff'],
      },
      { key: 'collect', name: '2 weeks to collect documents required for DP FMEA First Pass', days: 14, hours: 16 },
      { key: 'develop', name: 'DOC: DP FMEA First Pass (under development)', days: 14, hours: 1 },
      { key: 'qa', name: '4 weeks to Q&A and resolve received First Pass findings', days: 28, hours: 73 },
      note('stakeholders'),
      { key: 'vacation', name: '4 weeks Summer vacation', days: 28, vacation: true },
      ],
    ),
  );

  const rev1Head = 'rev1.upload';
  phases.push(
    backward('rev0', 'Full DP FMEA rev.0', rev1Head, [
      {
        key: 'collect',
        name: '2 weeks to collect documents required for full DP FMEA',
        days: 14,
        hours: 16,
        mustFollow: ['firstPass'],
      },
      {
        key: 'ioSegregation',
        name: '9.5.2. Check DPCS IOs segregation',
        days: 1,
        hours: 1,
        added: true,
        anchor: { kind: 'with', ref: 'rev0.collect' },
      },
      { key: 'cmc', name: "Review and update suppliers' CMC approval status", days: 14 },
      {
        key: 'r0',
        name: 'DOC: DP FMEA - R.0 (DP2 12 weeks / DP3 12-14 weeks for maker to produce DP FMEA)',
        days: 56,
        hours: 1,
        emphasis: true,
      },
      {
        key: 'cmcDuring',
        name: "Review and update suppliers' CMC approval status",
        days: 21,
        hours: 14,
        anchor: { kind: 'with', ref: 'rev0.r0' },
      },
      { key: 'qa', name: '3 weeks QA the full DP FMEA by all departments', days: 21, hours: 8 },
      { key: 'resolve', name: '4 weeks to resolve received findings', days: 28, hours: 68 },
      { key: 'r1', name: 'DOC: DP FMEA - R.1 (1 week after all findings are resolved)', days: 7, hours: 1, emphasis: true },
      { key: 'qaR1', name: 'Perform QA on DP FMEA update from R.0 to R.1', days: 4, hours: 6 },
      { key: 'inform', name: 'Inform VP PM that DP FMEA is ready for Owner & Class review', days: 1, hours: 1 },
    ]),
  );

  if (cfg.ownerReview) {
    phases.push(
      forward('owner', 'DP FMEA Owner review rev.1 (where applicable)', { kind: 'with', ref: rev1Head }, [
        { key: 'upload', name: `${UPLOAD} for Owners review`, days: 2, hours: 1 },
        { key: 'review', name: "2 weeks Owner's review", days: 14, hours: 1 },
        { key: 'respond', name: '2 weeks for VARD to respond comments & update & QC', days: 14, hours: 33 },
        note('stakeholders'),
        { key: 'vacation', name: '4 weeks summer vacation', days: 28, vacation: true },
      ]),
    );
  }

  for (let n = 1; n <= revs; n++) {
    const isLast = n === revs;
    const tail = isLast ? 'trials.delivery' : `rev${n + 1}.upload`;
    const phase = backward(`rev${n}`, `DP FMEA Class approval rev.${n}`, tail, classRevisionSpecs(n, isLast));
    if (!cfg.ownerReview) {
      for (const t of phase.tasks) t.mustFollow = t.mustFollow?.filter((r) => !r.startsWith('owner.'));
    }
    phases.push(phase);
  }

  const sea = 'trials.seaTrial';
  phases.push(
    forward('trials', 'DP FMEA Proving trials', { kind: 'date', key: 'seaTrial' }, [
      {
        key: 'delivery',
        name: 'DP FMEA delivery goal (10w before Sea trial)',
        days: 70,
        hours: 1,
        anchor: { kind: 'before', ref: sea },
      },
      {
        key: 'skills',
        name: 'Prepare required skills matrix and inform suppliers of needs (6w before Sea trial)',
        days: 42,
        hours: 1,
        anchor: { kind: 'before', ref: sea },
      },
      {
        key: 'skillsDoc',
        name: 'DOC: DP FMEA seatrial skills matrix',
        days: 3,
        hours: 7,
        anchor: { kind: 'with', ref: 'trials.skills' },
      },
      {
        key: 'coordination',
        name: 'Organize pre seatrial coordination meeting (4w before Sea trial)',
        days: 28,
        hours: 1,
        anchor: { kind: 'before', ref: sea },
      },
      {
        key: 'coordMeeting',
        name: 'Pre seatrial coordination meeting',
        days: 1,
        hours: 8,
        anchor: { kind: 'with', ref: 'trials.coordination' },
      },
      {
        key: 'coordMom',
        name: 'MOM: DP FMEA pre seatrial coordination',
        days: 1,
        hours: 2,
        anchor: { kind: 'after', ref: 'trials.coordMeeting' },
      },
      {
        key: 'labels',
        name: 'Order DP "FMEA" red labels to the workshop',
        days: 1,
        hours: 1,
        anchor: { kind: 'after', ref: 'trials.coordMom' },
      },
      {
        key: 'internalTesting',
        name: 'DP FMEA internal testing (2 weeks before sea trials)',
        days: 14,
        hours: 24,
        emphasis: true,
        anchor: { kind: 'before', ref: sea },
      },
      { key: 'seaTrial', name: 'Sea trial', days: 6, hours: 1, emphasis: true },
      { key: 'prelim', name: 'DOC: Preliminary DP FMEA proving trials findings', days: 1, hours: 1 },
      { key: 'teamWorks', name: 'DP FMEA team works (closing of findings)', days: 7, hours: 35 },
      { key: 'distribute', name: 'Distribute findings', days: 1 },
      { key: 'responses', name: 'Collect responses to findings', days: 1 },
      { key: 'documents', name: 'Collect updated documents', days: 1 },
      { key: 'meetings', name: 'Arrange for meetings with necessary stakeholders', days: 1 },
      { key: 'submit', name: 'Submit responses and updated docs to DP FMEA maker', days: 1 },
      {
        key: 'report',
        name: 'DOC: DP FMEA compiled report received',
        days: 7,
        hours: 1,
        emphasis: true,
        anchor: { kind: 'after', ref: 'trials.teamWorks' },
      },
      {
        key: 'reportUpload',
        name: `${UPLOAD} DP FMEA report for Class for information`,
        days: 1,
        hours: 1,
        anchor: { kind: 'after', ref: 'trials.report' },
      },
    ]),
  );

  phases.push(
    forward('lessons', 'DP FMEA lesson learned', { kind: 'after', ref: 'trials.reportUpload' }, [
      { key: 'lessons', name: 'Lesson learned (to be done before closure)', days: 14, hours: 14 },
    ]),
  );

  phases.push(
    forward('closure', 'DP FMEA project closure', { kind: 'date', key: 'vesselDelivery', offsetMonths: 3 }, [
      { key: 'closure', name: 'Project closure (3 months after vessel delivery)', days: 0, kind: 'milestone' },
    ]),
  );

  return phases;
}

/** Default key dates = the ones in the Excel template, so the first screen matches it. */
export const EXCEL_DEFAULT_DATES: {
  projectStart: string;
  firstPassStart: string | null;
  seaTrial: string;
  vesselDelivery: string | null;
} = {
  projectStart: '2024-01-02',
  firstPassStart: '2024-05-08',
  seaTrial: '2026-08-04',
  vesselDelivery: null,
};

/** Settings that make the planner calculate exactly like the Excel sheet. */
export const EXCEL_MODE = { calendar: 'none', workdayStarts: false, vacations: 'always' } as const;
/** Recommended settings: Norwegian calendar, no weekend starts, vacation only in summer. */
export const REALISTIC_MODE = { calendar: 'NO', workdayStarts: true, vacations: 'auto' } as const;
