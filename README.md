# DP FMEA Planner

Web replacement for the *DP FMEA process schedule* Excel Gantt template. Type the sea trial date and the
whole DP FMEA delivery plan is recalculated; change any date or duration and everything that depends on it
moves with it.

```bash
npm install
npm run dev        # http://localhost:5173
npm test           # scheduling engine tests (incl. comparison with the Excel template)
npm run build      # static site in dist/
```

## How the Excel template worked (and how this app models it)

The spreadsheet had two chains of formulas that meet in the middle:

| Chain | Driven by | Phases | Excel formula pattern |
|---|---|---|---|
| Forward | Project start (`C8`) | Preparation & planning → RFQ → Kick-off meetings | `START = previous END + 1` |
| Forward | First Pass start (typed in `C39`) | DP FMEA First Pass | `START = previous END + 1` |
| **Backward** | **Sea trial (`C90`)** | Proving trials ← Class approval rev.2 ← rev.1 ← Full DP FMEA rev.0 | `phase START = next phase START − SUM(DAYS)` |
| Parallel | Class approval rev.1 start | Owner review rev.1 | `START = C63` |

The app keeps exactly these rules, but as explicit links per task instead of cell formulas
(`src/engine/template.ts`):

- `after` – starts the day after its predecessor ends (forward planning)
- `before` – ends the day before its successor starts (backward planning from the sea trial)
- `with` – starts together with another task (parallel work, e.g. CMC review during R.0)
- `date` – starts on a key date (project start, First Pass start, sea trial, vessel delivery + 3 months)

Durations are calendar days, like the DAYS column; WORK DAYS is `NETWORKDAYS` (Mon–Fri).

### Verified against the spreadsheet

In **Excel-identical** mode the app reproduces the sheet exactly. The template's own formulas were
recalculated for four key-date scenarios (different sea trial, project start and First Pass dates) and
every START, END and WORK DAYS value of all 75 dated rows matches. The same holds for a real project file
(*prototype.xlsx*: project start 2026-09-21, sea trial 2029-01-31) using the values Excel itself saved
(`src/engine/excel-scenarios.test.ts`,
fixture in `src/engine/fixtures/excel-scenarios.json`). One deliberate difference: Excel adds the parallel task
*9.5.2 Check DPCS IOs segregation* into the rev.0 phase duration (`=SUM(F47:F56)-F51`), which starts rev.0
one day too early and leaves a one-day gap before Class approval rev.1. The app does not copy that.

### Realistic mode (default) – where the Excel logic gives nonsense

The sheet's arithmetic is right, but three of its rules produce plans no one would follow. **Realistic**
mode (the default; switch in *Scope → Calculation*) fixes them:

| Excel rule | Problem | Realistic mode |
|---|---|---|
| First Pass start is typed by hand | Change the project dates and First Pass stays in the old year (prototype.xlsx: 2024, two years before project start) | Leave it empty and First Pass starts after the kick-off meetings; typing a date still works |
| Three fixed "4 weeks summer vacation" rows in the chain | They land wherever the chain puts them (prototype.xlsx: Aug–Sep 2028) while the real July vacation hits "resolve findings" with nothing counted | Company vacations are calendar periods – **summer from 1 July, 28 days** and **Christmas from 22 December, 14 days** (both editable in *Scope*). They appear as rows in the phases they interrupt, and work pauses during them (a 28-day task through July ends 28 days later). Short tasks (≤ 5 days: meetings, uploads) move past a vacation instead of being split |
| Calendar-day arithmetic | Uploads, meetings and DOC tasks start on Saturdays, Sundays and holidays (0 work days) | Tasks start on working days (no weekends, Norwegian public holidays or company vacation); forward-planned tasks move later, backward-planned tasks earlier. Key dates and pinned dates are never moved |

## What you can do that Excel made painful

- **Key dates**: sea trial, project start, First Pass start, vessel delivery. The result panel shows when
  rev.0 must start, the final Class approval, the DP FMEA delivery goal, and the **buffer** between the end of
  First Pass and the start of rev.0 (red when the plan no longer fits).
- **Class approval revisions**: `+`/`−` adds *DP FMEA Class approval rev.3, rev.4 …* (up to 9). Each extra
  cycle is *upload → 1-3 weeks Class approval → respond & QC → DOC R.N+1 → QA*. The final
  *upload + Class approval* stays at the end of the last revision, so everything before it moves earlier.
- **Change any date**: typing a start or end date pins the task (📌); its dependants follow. Pins that break
  the logic (e.g. starting before the predecessor has finished) are listed as conflicts.
- **Change any duration / hours / name**, or **skip** a task (e.g. no summer vacation): the chain closes up.
- **Drag** a bar to move it, drag its right edge to change the duration.
- **Norwegian calendar**: work days exclude Norwegian public holidays (the same *helligdager* as
  [norskkalender.no](https://www.norskkalender.no/), computed from the Easter date so it works for any
  year). Holidays are shaded in the Gantt, and meetings, uploads, deliverables, the sea trial and pinned dates
  that start on a holiday are flagged. Fellesferie (ISO weeks 28–30) can be shown as well. Switch the calendar
  off to get Excel's plain `NETWORKDAYS` (Mon–Fri).
- **Owner review** on/off, checklist notes on/off, **Collapse all / Expand all** (remembered per project).
- **Baseline**: freeze today's plan and see how many days each task moved later.
- **Undo/redo** (Ctrl+Z / Ctrl+Shift+Z), several projects, duplicate a project to try a scenario.
- **Export Excel** (dates + weekly Gantt) for people who still want the spreadsheet, **Save/Open file**
  (`.dpfmea.json`) to share a project.

Data is stored in the browser (localStorage) – use *Save file* to share or back up.

## Stack

| Concern | Choice | Why |
|---|---|---|
| UI | **React 19** | Standard, big talent pool, fits a table + Gantt UI well |
| Language | **TypeScript 7** (native compiler) | `strict` by default, ~10× faster type-checking; no Vue/Svelte, so no TS 6 tooling constraints |
| Build/dev | **Vite 8** | Instant dev server, static output that can be hosted anywhere |
| Tests | **Vitest** | Same config as Vite; the scheduling engine is pure TypeScript and fully unit-tested |
| Dates | Own whole-day arithmetic (`src/engine/dates.ts`) | UTC day numbers → no time-zone/DST bugs, no date library needed |
| Gantt | Own lightweight renderer | ~100 rows need no heavy Gantt library; full control over backward/forward logic and VARD styling |
| Excel export | **exceljs** (lazy-loaded) | Only downloaded when someone clicks *Export Excel* |
| State | `useReducer` + localStorage | Undo/redo for free; no server needed for v1 |

The scheduling engine (`src/engine/`) has no React dependency, so it can later run on a server unchanged.

### Suggested next steps

1. **Hosting**: the app is published to GitHub Pages by `.github/workflows/pages.yml` on every push to the
   default branch (https://yuriiss.github.io/vard/). Any static host works too (Azure Static Web Apps,
   an internal web server).
2. **Shared projects**: add a small API + database (e.g. Node/Fastify or ASP.NET + PostgreSQL, or Supabase)
   with Microsoft Entra ID sign-in, storing the same `Project` JSON the app already saves.
3. Optional "move to next working day" for meetings/uploads that land on a weekend or holiday.
4. **Custom tasks** per project and template versioning (the Excel template has been updated over time).
5. **Import** of existing project Excel files (read the key dates from `C8`, `C39`, `C90`).

## Project layout

```
src/engine/dates.ts      whole-day date arithmetic
src/engine/template.ts   the DP FMEA process (phases, tasks, links) – edit here to change the standard template
src/engine/scheduler.ts  calculates dates, phase spans, conflicts and buffer
src/state/store.ts       projects, edits, undo/redo, localStorage
src/state/files.ts       Excel export, JSON save/open
src/components/          UI (key dates, schedule table + Gantt)
```
