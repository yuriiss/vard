// Cross-check against the Excel template itself: the sheet's own formulas were recalculated
// for several key-date scenarios (see fixtures/excel-scenarios.json) and every START, END and
// WORK DAYS value must match the planner.
import { describe, expect, it } from 'vitest';
import data from './fixtures/excel-scenarios.json';
import { schedule } from './scheduler';
import { parseISO, toISO } from './dates';
import { EXCEL_MODE } from './template';

// Known Excel bug, not reproduced: the rev.0 phase duration includes the parallel task
// "9.5.2 Check DPCS IOs segregation", so Excel starts all of rev.0 one day early.
const shiftFor = (id: string) => (id.startsWith('rev0.') ? 1 : 0);

for (const [name, sc] of Object.entries(data.scenarios)) {
  describe(`Excel scenario "${name}" (sea trial ${sc.inputs.seaTrial})`, () => {
    const s = schedule({
      keyDates: { ...sc.inputs, vesselDelivery: null },
      classRevisions: 2,
      ownerReview: true,
      ...EXCEL_MODE, // plain NETWORKDAYS, calendar-day starts, vacation blocks always 28 days
      overrides: {},
    });
    for (const [id, x] of Object.entries(sc.rows)) {
      it(`row ${x.row} ${id}`, () => {
        const t = s.tasks.get(id)!;
        expect(t, id).toBeDefined();
        const k = shiftFor(id);
        expect(toISO(t.start!)).toBe(toISO(parseISO(x.start!) + k));
        expect(toISO(t.end!)).toBe(toISO(parseISO(x.end!) + k));
        if (x.workDays != null && k === 0) expect(t.workDays).toBe(x.workDays);
      });
    }
  });
}
