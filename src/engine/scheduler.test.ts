import { describe, expect, it } from 'vitest';
import { DEFAULT_SUMMER, DEFAULT_WINTER, schedule } from './scheduler';
import { fellesferie, vacationWindows, yearOf, isWeekend, holidayName, norwegianHolidays, toISO, parseISO, workDays } from './dates';
import { EXCEL_DEFAULT_DATES, REALISTIC_MODE } from './template';
import type { ProjectConfig } from './types';

const base = (): ProjectConfig => ({
  keyDates: { ...EXCEL_DEFAULT_DATES },
  classRevisions: 2,
  ownerReview: true,
  // Excel arithmetic, plus the Norwegian calendar for work days / warnings.
  calendar: 'NO',
  workdayStarts: false,
  vacations: 'always',
  overrides: {},
});

// START/END values Excel calculated for the template (sea trial 2026-08-04).
const EXCEL: [id: string, start: string, end: string][] = [
  ['start.period', '2024-01-02', '2024-01-15'], // row 8
  ['prep.exchange', '2024-01-16', '2024-01-16'], // row 10
  ['prep.plmStructure', '2024-01-17', '2024-01-17'], // row 11
  ['prep.plmPlan', '2024-01-18', '2024-01-18'], // row 12
  ['prep.docList', '2024-01-19', '2024-01-22'], // row 13
  ['prep.dependencies', '2024-01-23', '2024-01-28'], // row 14
  ['prep.kickoff', '2024-01-29', '2024-01-30'], // row 15
  ['prep.mom', '2024-01-31', '2024-01-31'], // row 20
  ['rfq.prepare', '2024-02-01', '2024-02-02'], // row 22
  ['rfq.submit', '2024-02-03', '2024-02-03'], // row 23
  ['rfq.wait', '2024-02-04', '2024-02-17'], // row 24
  ['rfq.evaluate', '2024-02-18', '2024-02-22'], // row 25
  ['rfq.nominate', '2024-02-23', '2024-02-23'], // row 26
  ['rfq.consult', '2024-02-24', '2024-02-24'], // row 27
  ['rfq.tss', '2024-02-25', '2024-02-27'], // row 28
  ['rfq.po', '2024-02-28', '2024-03-01'], // row 29
  ['rfq.waitPo', '2024-03-02', '2024-03-08'], // row 30
  ['kickoff.meeting', '2024-03-09', '2024-03-09'], // row 32
  ['kickoff.mom', '2024-03-10', '2024-03-12'], // row 33
  ['kickoff.matrix', '2024-03-13', '2024-03-15'], // row 34
  ['kickoff.mobilise', '2024-03-16', '2024-03-24'], // row 35
  ['kickoff.classMeeting', '2024-03-25', '2024-03-26'], // row 36
  ['kickoff.classMom', '2024-03-27', '2024-03-27'], // row 37
  ['firstPass.mark', '2024-05-08', '2024-05-10'], // row 40
  ['firstPass.collect', '2024-05-11', '2024-05-24'], // row 41
  ['firstPass.develop', '2024-05-25', '2024-06-07'], // row 42
  ['firstPass.qa', '2024-06-08', '2024-07-05'], // row 43
  ['firstPass.vacation', '2024-07-06', '2024-08-02'], // row 45
  ['rev0.collect', '2025-09-02', '2025-09-15'], // row 47
  ['rev0.ioSegregation', '2025-09-02', '2025-09-02'], // row 48
  ['rev0.cmc', '2025-09-16', '2025-09-29'], // row 49
  ['rev0.r0', '2025-09-30', '2025-11-24'], // row 50
  ['rev0.cmcDuring', '2025-09-30', '2025-10-20'], // row 51
  ['rev0.qa', '2025-11-25', '2025-12-15'], // row 52
  ['rev0.resolve', '2025-12-16', '2026-01-12'], // row 53
  ['rev0.r1', '2026-01-13', '2026-01-19'], // row 54
  ['rev0.qaR1', '2026-01-20', '2026-01-23'], // row 55
  ['rev0.inform', '2026-01-24', '2026-01-24'], // row 56
  ['owner.upload', '2026-01-26', '2026-01-27'], // row 58
  ['owner.review', '2026-01-28', '2026-02-10'], // row 59
  ['owner.respond', '2026-02-11', '2026-02-24'], // row 60
  ['owner.vacation', '2026-02-25', '2026-03-24'], // row 62
  ['rev1.upload', '2026-01-26', '2026-01-26'], // row 64
  ['rev1.approval', '2026-01-27', '2026-02-16'], // row 65
  ['rev1.respond', '2026-02-17', '2026-03-02'], // row 66
  ['rev1.vacation', '2026-03-03', '2026-03-30'], // row 68
  ['rev1.doc', '2026-03-31', '2026-04-02'], // row 69
  ['rev1.qa', '2026-04-03', '2026-04-04'], // row 70
  ['rev2.upload', '2026-04-05', '2026-04-07'], // row 72
  ['rev2.testProcedure', '2026-04-08', '2026-04-08'], // row 73
  ['rev2.approval', '2026-04-08', '2026-04-21'], // row 74
  ['rev2.respond', '2026-04-22', '2026-05-05'], // row 75
  ['rev2.doc', '2026-05-06', '2026-05-07'], // row 77
  ['rev2.qa', '2026-05-08', '2026-05-10'], // row 78
  ['classFinal.upload', '2026-05-11', '2026-05-11'], // row 79
  ['classFinal.approval', '2026-05-12', '2026-05-25'], // row 80
  ['trials.delivery', '2026-05-26', '2026-08-03'], // row 82
  ['trials.skills', '2026-06-23', '2026-08-03'], // row 83
  ['trials.skillsDoc', '2026-06-23', '2026-06-25'], // row 84
  ['trials.coordination', '2026-07-07', '2026-08-03'], // row 85
  ['trials.coordMeeting', '2026-07-07', '2026-07-07'], // row 86
  ['trials.coordMom', '2026-07-08', '2026-07-08'], // row 87
  ['trials.labels', '2026-07-09', '2026-07-09'], // row 88
  ['trials.internalTesting', '2026-07-21', '2026-08-03'], // row 89
  ['trials.seaTrial', '2026-08-04', '2026-08-09'], // row 90
  ['trials.prelim', '2026-08-10', '2026-08-10'], // row 91
  ['trials.teamWorks', '2026-08-11', '2026-08-17'], // row 92
  ['trials.distribute', '2026-08-18', '2026-08-18'], // row 93
  ['trials.responses', '2026-08-19', '2026-08-19'], // row 94
  ['trials.documents', '2026-08-20', '2026-08-20'], // row 95
  ['trials.meetings', '2026-08-21', '2026-08-21'], // row 96
  ['trials.submit', '2026-08-22', '2026-08-22'], // row 97
  ['trials.report', '2026-08-18', '2026-08-24'], // row 98
  ['trials.reportUpload', '2026-08-25', '2026-08-25'], // row 99
  ['lessons.lessons', '2026-08-26', '2026-09-08'], // row 101
];

// Excel sums "9.5.2. Check DPCS IOs segregation" (a parallel task) into the rev.0 phase
// duration, so the whole rev.0 block starts one day too early and leaves a 1-day gap
// before class approval rev.1. The planner does not reproduce that bug.
const EXCEL_REV0_BUG_SHIFT = 1;

const iso = (d: number | null) => (d == null ? null : toISO(d));

describe('matches the Excel template', () => {
  const s = schedule(base());
  for (const [id, start, end] of EXCEL) {
    it(id, () => {
      const t = s.tasks.get(id);
      expect(t, id).toBeDefined();
      const shift = id.startsWith('rev0.') ? EXCEL_REV0_BUG_SHIFT : 0;
      expect(iso(t!.start)).toBe(toISO(parseISO(start) + shift));
      expect(iso(t!.end)).toBe(toISO(parseISO(end) + shift));
    });
  }

  it('has no conflicts for the template dates', () => {
    expect(s.warnings.filter((w) => w.severity === 'error')).toEqual([]);
  });

  it('flags the Easter Sunday upload in the template (Norwegian calendar)', () => {
    expect(s.warnings).toEqual([
      { taskId: 'rev2.upload', severity: 'warning', message: 'Starts on a public holiday: 1. påskedag (05 Apr 2026).' },
    ]);
    const plain = schedule({ ...base(), calendar: 'none' });
    expect(plain.warnings).toEqual([]);
  });

  it('reports the slack between First Pass and rev.0', () => {
    // First Pass ends 2024-08-02, rev.0 starts 2025-09-03.
    expect(s.buffer).toBe(parseISO('2025-09-03') - parseISO('2024-08-02') - 1);
  });
});

describe('sea trial drives the backward chain', () => {
  it('moves every backward-planned task by the same amount', () => {
    const a = schedule(base());
    const cfg = base();
    cfg.keyDates.seaTrial = '2026-09-01'; // +28 days
    const b = schedule(cfg);
    for (const id of ['rev0.collect', 'rev1.upload', 'owner.review', 'classFinal.approval', 'trials.delivery']) {
      expect(b.tasks.get(id)!.start! - a.tasks.get(id)!.start!, id).toBe(28);
    }
    // Forward chain is untouched.
    expect(b.tasks.get('rfq.po')!.start).toBe(a.tasks.get('rfq.po')!.start);
  });

  it('flags the collision when the sea trial is too early', () => {
    const cfg = base();
    cfg.keyDates.seaTrial = '2025-03-01';
    const s = schedule(cfg);
    expect(s.buffer!).toBeLessThan(0);
    expect(s.warnings.some((w) => w.taskId === 'rev0.collect' && w.severity === 'error')).toBe(true);
  });
});

describe('class approval revisions', () => {
  it('rev.3 is inserted before the final approval and pushes earlier work back', () => {
    const two = schedule(base());
    const cfg = base();
    cfg.classRevisions = 3;
    const three = schedule(cfg);

    expect(three.phases.map((p) => p.id)).toContain('rev3');
    const rev3 = three.phases.find((p) => p.id === 'rev3')!;
    expect(rev3.name).toBe('DP FMEA Class approval rev.3');
    expect(rev3.tasks.map((t) => t.id)).toContain('classFinal.approval');
    expect(three.phases.find((p) => p.id === 'rev2')!.tasks.map((t) => t.id)).not.toContain('classFinal.approval');
    expect(rev3.tasks.some((t) => t.name.startsWith('DOC: DP FMEA - R.4'))).toBe(true);

    // Final approval still ends the day before the delivery goal.
    expect(three.tasks.get('classFinal.approval')!.end).toBe(two.tasks.get('classFinal.approval')!.end);

    // rev.3 adds upload(3) + approval(14) + respond(14) + doc(2) + QA(3) = 36 days.
    expect(two.tasks.get('rev1.upload')!.start! - three.tasks.get('rev1.upload')!.start!).toBe(36);
    expect(three.warnings.filter((w) => w.severity === 'error')).toEqual([]);
  });

  it('supports a single revision', () => {
    const cfg = base();
    cfg.classRevisions = 1;
    const s = schedule(cfg);
    expect(s.phases.some((p) => p.id === 'rev2')).toBe(false);
    expect(s.tasks.get('rev1.qa')!.end! + 1).toBe(s.tasks.get('classFinal.upload')!.start);
  });

  it('keeps edits on the final approval when revisions are added', () => {
    const cfg = base();
    cfg.overrides['classFinal.approval'] = { duration: 21 };
    cfg.classRevisions = 4;
    const s = schedule(cfg);
    expect(s.tasks.get('classFinal.approval')!.days).toBe(21);
    expect(s.tasks.get('classFinal.approval')!.phaseId).toBe('rev4');
  });
});

describe('overrides', () => {
  it('pinning a start date moves its successors', () => {
    const cfg = base();
    cfg.overrides['rfq.wait'] = { pinStart: '2024-02-10' }; // was 2024-02-04
    const s = schedule(cfg);
    expect(iso(s.tasks.get('rfq.evaluate')!.start)).toBe('2024-02-24');
    expect(s.warnings.find((w) => w.taskId === 'rfq.wait')?.severity).toBe('warning'); // gap
  });

  it('pinning inside the backward chain moves its predecessors', () => {
    const cfg = base();
    cfg.overrides['rev1.approval'] = { pinStart: '2026-01-20' }; // was 2026-01-27
    const s = schedule(cfg);
    expect(iso(s.tasks.get('rev1.upload')!.start)).toBe('2026-01-19');
    expect(iso(s.tasks.get('rev0.inform')!.end)).toBe('2026-01-18');
    expect(s.warnings.some((w) => w.taskId === 'rev1.approval')).toBe(false);
  });

  it('pinning an end date keeps the duration', () => {
    const cfg = base();
    cfg.overrides['rfq.evaluate'] = { pinEnd: '2024-03-01' };
    const s = schedule(cfg);
    expect(iso(s.tasks.get('rfq.evaluate')!.start)).toBe('2024-02-26');
  });

  it('a pin that overlaps its predecessor is an error', () => {
    const cfg = base();
    cfg.overrides['rfq.evaluate'] = { pinStart: '2024-02-10' };
    const s = schedule(cfg);
    expect(s.warnings.find((w) => w.taskId === 'rfq.evaluate')?.severity).toBe('error');
  });

  it('disabling a vacation shortens the chain', () => {
    const cfg = base();
    cfg.overrides['rev1.vacation'] = { disabled: true };
    const s = schedule(cfg);
    const a = schedule(base());
    expect(s.tasks.get('rev1.upload')!.start! - a.tasks.get('rev1.upload')!.start!).toBe(28);
  });

  it('changing a duration ripples through', () => {
    const cfg = base();
    cfg.overrides['rev0.r0'] = { duration: 84 }; // 12 weeks instead of 8
    const s = schedule(cfg);
    const a = schedule(base());
    expect(a.tasks.get('rev0.collect')!.start! - s.tasks.get('rev0.collect')!.start!).toBe(28);
  });

  it('closure is 3 months after vessel delivery', () => {
    const cfg = base();
    expect(schedule(cfg).tasks.get('closure.closure')!.start).toBeNull();
    cfg.keyDates.vesselDelivery = '2026-11-30';
    expect(iso(schedule(cfg).tasks.get('closure.closure')!.start)).toBe('2027-02-28');
  });

  it('owner review can be left out', () => {
    const cfg = base();
    cfg.ownerReview = false;
    const s = schedule(cfg);
    expect(s.tasks.has('owner.upload')).toBe(false);
    expect(s.warnings.filter((w) => w.severity === 'error')).toEqual([]);
  });
});

describe('Norwegian calendar (norskkalender.no)', () => {
  // "Helligdager" as listed on https://www.norskkalender.no/?year=YYYY
  const SITE: Record<number, string[]> = {
    2025: ['2025-01-01', '2025-04-13', '2025-04-17', '2025-04-18', '2025-04-20', '2025-04-21', '2025-05-01', '2025-05-17', '2025-05-29', '2025-06-08', '2025-06-09', '2025-12-25', '2025-12-26'],
    2026: ['2026-01-01', '2026-03-29', '2026-04-02', '2026-04-03', '2026-04-05', '2026-04-06', '2026-05-01', '2026-05-14', '2026-05-17', '2026-05-24', '2026-05-25', '2026-12-25', '2026-12-26'],
    // 17 May 2027 is also 2. pinsedag – one day, two names.
    2027: ['2027-01-01', '2027-03-21', '2027-03-25', '2027-03-26', '2027-03-28', '2027-03-29', '2027-05-01', '2027-05-06', '2027-05-16', '2027-05-17', '2027-12-25', '2027-12-26'],
  };
  for (const [year, days] of Object.entries(SITE)) {
    it(`matches ${year}`, () => {
      expect(norwegianHolidays(Number(year)).map((h) => toISO(h.day))).toEqual(days);
    });
  }

  it('work days exclude weekday holidays', () => {
    // Easter week 2026: Mon 30 Mar – Mon 6 Apr = 6 weekdays, minus Thu, Fri, Mon holidays.
    expect(workDays(parseISO('2026-03-30'), parseISO('2026-04-06'))).toBe(6);
    expect(workDays(parseISO('2026-03-30'), parseISO('2026-04-06'), true)).toBe(3);
  });

  it('fellesferie is ISO weeks 28–30', () => {
    expect(fellesferie(2026).map(toISO)).toEqual(['2026-07-06', '2026-07-26']);
  });
});

describe('realistic mode', () => {
  const realistic = (kd: Partial<ProjectConfig['keyDates']> = {}): ProjectConfig => ({
    ...base(),
    ...REALISTIC_MODE,
    keyDates: { ...EXCEL_DEFAULT_DATES, ...kd },
  });

  it('First Pass follows the kick-off meetings when no date is typed', () => {
    const s = schedule(realistic({ projectStart: '2026-10-05', firstPassStart: null, seaTrial: '2028-06-30' }));
    const kickoffEnd = s.tasks.get('kickoff.classMom')!.end!;
    const fp = s.tasks.get('firstPass.mark')!.start!;
    expect(fp).toBeGreaterThan(kickoffEnd);
    expect(fp - kickoffEnd).toBeLessThanOrEqual(4); // next working day (Christmas/weekend at most)
    expect(s.warnings.filter((w) => w.severity === 'error')).toEqual([]);
  });

  it('no task that follows another starts on a weekend or Norwegian holiday', () => {
    for (const seaTrial of ['2026-08-04', '2027-03-15', '2028-06-30']) {
      const s = schedule(realistic({ firstPassStart: null, seaTrial }));
      for (const t of s.tasks.values()) {
        if (t.start == null || t.days === 0 || t.anchor.kind === 'date' || t.anchor.kind === 'none') continue;
        if (t.calendarBlock) continue;
        expect(isWeekend(t.start) || !!holidayName(t.start), `${seaTrial} ${t.id} ${toISO(t.start)}`).toBe(false);
      }
    }
  });

  it('backward-planned tasks still finish before their successor starts', () => {
    const s = schedule(realistic());
    for (const t of s.tasks.values()) {
      if (t.anchor.kind !== 'before' || t.start == null || t.days === 0) continue;
      const succ = s.tasks.get(t.anchor.ref);
      if (!succ || succ.start == null) continue; // hidden 0-day vacation row
      expect(t.end!, t.id).toBeLessThan(succ.start);
    }
  });

  const seaTrials = () => {
    const out: string[] = [];
    // Sea trial every second Monday over three years.
    for (let d = parseISO('2026-01-05'); d < parseISO('2029-01-05'); d += 14) out.push(toISO(d));
    return out;
  };
  const windows = (s: ReturnType<typeof schedule>) =>
    vacationWindows(yearOf(s.start!), yearOf(s.end!), DEFAULT_SUMMER, DEFAULT_WINTER);

  it('work never starts or ends inside a company vacation', () => {
    for (const seaTrial of seaTrials()) {
      const s = schedule(realistic({ firstPassStart: null, projectStart: '2025-06-02', seaTrial }));
      const ws = windows(s);
      for (const t of s.tasks.values()) {
        if (t.calendarBlock || t.start == null || t.days === 0 || t.anchor.kind === 'date') continue;
        if (t.anchor.kind === 'before' && t.anchor.ref === 'trials.seaTrial') continue;
        for (const v of ws) {
          expect(t.start >= v.start && t.start <= v.end, `${seaTrial} ${t.id} starts in ${v.name}`).toBe(false);
          expect(t.end! >= v.start && t.end! <= v.end, `${seaTrial} ${t.id} ends in ${v.name}`).toBe(false);
        }
      }
    }
  });

  it('a task that runs through a vacation is stretched by exactly the vacation days', () => {
    for (const seaTrial of seaTrials()) {
      const s = schedule(realistic({ seaTrial }));
      const ws = windows(s);
      for (const t of s.tasks.values()) {
        if (t.calendarBlock || t.start == null || t.days === 0) continue;
        const overlap = ws.reduce((n, v) => n + Math.max(0, Math.min(t.end!, v.end) - Math.max(t.start!, v.start) + 1), 0);
        if (t.pauseDays > 0) expect(t.pauseDays, `${seaTrial} ${t.id}`).toBe(overlap);
        expect(t.end! - t.start! + 1, `${seaTrial} ${t.id}`).toBe(t.days + t.pauseDays);
      }
    }
  });

  it('every vacation that pauses work is shown as a row in that phase', () => {
    for (const seaTrial of seaTrials()) {
      const s = schedule(realistic({ seaTrial }));
      for (const p of s.phases) {
        for (const t of p.tasks) {
          if (!t.pauseDays) continue;
          const rows = p.tasks.filter((r) => r.calendarBlock && r.start! <= t.end! && r.end! >= t.start!);
          expect(rows.length, `${seaTrial} ${t.id}`).toBeGreaterThan(0);
        }
      }
      // The template's fixed "4 weeks summer vacation" rows are replaced by the calendar ones.
      expect(s.tasks.has('rev1.vacation')).toBe(false);
    }
  });

  it('short tasks (meetings, uploads) are moved past a vacation, never split', () => {
    for (const seaTrial of seaTrials()) {
      const s = schedule(realistic({ firstPassStart: null, projectStart: '2025-06-02', seaTrial }));
      for (const t of s.tasks.values()) if (t.days > 0 && t.days <= 5) expect(t.pauseDays, `${seaTrial} ${t.id}`).toBe(0);
    }
  });

  it('real project (prototype.xlsx): summer 2028 pauses rev.0, Christmas shown before First Pass', () => {
    const s = schedule(realistic({ projectStart: '2026-09-21', firstPassStart: null, seaTrial: '2029-01-31' }));
    const rev0 = s.phases.find((p) => p.id === 'rev0')!;
    const summer = rev0.tasks.find((t) => t.calendarBlock && t.name.startsWith('Summer vacation 2028'));
    expect(summer, 'summer 2028 row in rev.0').toBeDefined();
    expect(toISO(summer!.start!)).toBe('2028-07-01');
    expect(toISO(summer!.end!)).toBe('2028-07-28');
    expect(s.tasks.get('rev0.resolve')!.pauseDays).toBe(28);
    const xmas = [...s.tasks.values()].find((t) => t.calendarBlock && t.name.startsWith('Christmas vacation 2026/27'));
    expect(xmas, 'Christmas 2026/27 row').toBeDefined();
    // The Class kick-off meeting is not split over Christmas; it moves to after it.
    const meeting = s.tasks.get('kickoff.classMeeting')!;
    expect(meeting.pauseDays).toBe(0);
    expect(meeting.start!).toBeGreaterThan(xmas!.end!);
    expect(s.warnings.filter((w) => w.severity === 'error')).toEqual([]);
  });

  it('Excel-identical mode keeps the fixed 28-day rows', () => {
    const s = schedule({ ...realistic(), vacations: 'always' });
    expect(s.tasks.get('rev1.vacation')!.days).toBe(28);
    expect([...s.tasks.values()].some((t) => t.calendarBlock)).toBe(false);
  });
});
