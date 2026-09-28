import { holidaysBetween, isWeekend, toDate, startOfWeek, isoWeek, type Day } from '../engine/dates';
import type { Project, Schedule } from '../engine/types';

export function download(filename: string, blob: Blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const safeName = (s: string) => s.replace(/[^\w.-]+/g, '_').replace(/^_+|_+$/g, '') || 'dp-fmea';

export function exportJson(project: Project) {
  download(`${safeName(project.name)}.dpfmea.json`, new Blob([JSON.stringify(project, null, 2)], { type: 'application/json' }));
}

export async function readJsonFile(file: File): Promise<unknown> {
  return JSON.parse(await file.text());
}

/** Excel export with the calculated dates and a weekly Gantt, for sharing with people who still live in Excel. */
export async function exportExcel(project: Project, s: Schedule) {
  const { default: ExcelJS } = await import('exceljs');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'DP FMEA Planner';
  const ws = wb.addWorksheet('DP FMEA schedule', {
    views: [{ state: 'frozen', xSplit: 7, ySplit: 7 }],
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  const RED = 'FFD20000';
  const INK = 'FF323232';
  const GREY = 'FF848484';
  const LIGHT = 'FFD6D6D6';

  ws.getCell('A1').value = 'DP FMEA process schedule';
  ws.getCell('A1').font = { name: 'Cambria', size: 16, bold: true };
  ws.getCell('A2').value = `${project.name}${project.vessel ? ` – ${project.vessel}` : ''}`;
  ws.getCell('A3').value = `Sea trial: ${project.keyDates.seaTrial}   Project start: ${project.keyDates.projectStart}   First Pass start: ${project.keyDates.firstPassStart}   Class approval revisions: ${project.classRevisions}`;
  ws.getCell('A4').value = `Exported ${new Date().toISOString().slice(0, 10)} from DP FMEA Planner`;
  ws.getCell('A4').font = { italic: true, color: { argb: GREY } };

  const header = ['WBS', 'TASK', 'START', 'END', 'DAYS', 'WORK DAYS', 'HOURS'];
  const first = s.start != null ? startOfWeek(s.start) : null;
  const weeks: Day[] = [];
  if (first != null && s.end != null) for (let d = first; d <= s.end; d += 7) weeks.push(d);

  const hr = ws.getRow(7);
  header.forEach((h, i) => (hr.getCell(i + 1).value = h));
  weeks.forEach((w, i) => {
    const c = ws.getRow(6).getCell(8 + i);
    c.value = toDate(w);
    c.numFmt = 'dd.mm.yy';
    c.alignment = { textRotation: 90 };
    hr.getCell(8 + i).value = `W${isoWeek(w)}`;
  });
  hr.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  hr.eachCell((c) => (c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: INK } }));

  ws.getColumn(1).width = 7;
  ws.getColumn(2).width = 70;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 12;
  for (let i = 5; i <= 7; i++) ws.getColumn(i).width = 8;
  for (let i = 0; i < weeks.length; i++) ws.getColumn(8 + i).width = 3.2;

  const paint = (rowNo: number, start: Day | null, end: Day | null, argb: string) => {
    if (start == null || end == null) return;
    weeks.forEach((w, i) => {
      if (start <= w + 6 && end >= w) {
        ws.getRow(rowNo).getCell(8 + i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
      }
    });
  };

  let r = 8;
  for (const p of s.phases) {
    const row = ws.getRow(r);
    row.values = [p.wbs, p.name, p.start != null ? toDate(p.start) : null, p.end != null ? toDate(p.end) : null, p.days, null, p.hours || null];
    row.font = { bold: true };
    for (let c = 1; c <= 7; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT } };
    paint(r, p.start, p.end, INK);
    r++;
    for (const t of p.tasks) {
      const row = ws.getRow(r);
      if (t.kind === 'note') {
        row.values = [null, `   • ${t.name}`];
        row.font = { italic: true, color: { argb: GREY } };
      } else {
        row.values = [
          t.wbs,
          `${t.name}${t.disabled ? ' (skipped)' : ''}${t.pinned ? ' 📌' : ''}`,
          t.start != null ? toDate(t.start) : null,
          t.end != null ? toDate(t.end) : null,
          t.days,
          t.workDays,
          t.hours ?? null,
        ];
        if (t.emphasis) row.font = { bold: true };
        if (t.disabled) row.font = { strike: true, color: { argb: GREY } };
        if (!t.disabled) paint(r, t.start, t.end, t.emphasis ? RED : t.vacation ? LIGHT : GREY);
      }
      r++;
    }
  }
  for (let i = 8; i < r; i++) {
    ws.getRow(i).getCell(3).numFmt = 'dd.mm.yyyy';
    ws.getRow(i).getCell(4).numFmt = 'dd.mm.yyyy';
  }

  if (project.calendar !== 'none' && s.start != null && s.end != null) {
    const hs = wb.addWorksheet('Holidays (NO)');
    hs.columns = [
      { header: 'Date', key: 'date', width: 12 },
      { header: 'Holiday', key: 'name', width: 34 },
      { header: 'Weekday', key: 'weekday', width: 10 },
    ];
    hs.getRow(1).font = { bold: true };
    for (const h of holidaysBetween(s.start, s.end)) {
      const row = hs.addRow({ date: toDate(h.day), name: h.name, weekday: isWeekend(h.day) ? 'weekend' : 'weekday' });
      row.getCell(1).numFmt = 'dd.mm.yyyy';
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  download(
    `${safeName(project.name)}.xlsx`,
    new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
  );
}
