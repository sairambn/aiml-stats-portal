# Result Analysis Portal

**Exam-cell ready** internal assessment result analysis.

Upload a section mark sheet → view diagrams and tables → download a professional **FRONT + MARK** Excel report.

**Live:** [aiml-stats-portal.vercel.app](https://aiml-stats-portal.vercel.app)

---

## Quick start (30 seconds)

1. Open the [live portal](https://aiml-stats-portal.vercel.app) **or** run locally (`npm install && npm run dev`).
2. Upload a mark sheet (`.xlsx`) — or click **Load demo data** to try instantly.
3. Review Overview diagrams, Subjects, Students, Toppers.
4. Optionally **Merge name list** (Reg.No + Name) for official spellings.
5. Adjust pass mark in the sidebar if needed (default **50**).
6. **Export analysis** → exam-cell Excel with borders, AB (yellow), fails (red).

No student data is stored on the server. Refresh clears the session.

---

## Sample file for testing

A ready-to-use sample is available:

| File | Contents |
|------|----------|
| `sample-mark-sheet-IA2.xlsx` | 30 students, 6 subjects (CS3452, CS3491, CS3492, IT3401, CS3451, GE3451), FRONT + MARK + NameList sheets |

Download from the repo (`public/samples/`) or use the built-in **Load demo data** button. Use the **NameList** sheet with **Merge name list** after upload to test name merging.

---

## What it does

| Step | Action |
|------|--------|
| 1 | **Upload mark sheet** (`.xlsx`) |
| 2 | Portal reads students, subjects, marks, absents |
| 3 | Overview **diagrams** + pass / fail / arrear stats |
| 4 | Optional **Merge name list** (official names by Reg.No) |
| 5 | **Export analysis** → exam-cell Excel with borders |

---

## Features

### Upload formats supported

- **Full RA sheet** — `FRONT` + `MARK` with B/G, C/M, H/DS, category tables  
- **Simple IAT / AIML mark sheet** — `S.No | Reg.No | Name | CS3452(TOC) | …`  
- Subject headers: `CS3452 & Theory of Computation` **or** `CS3452(TOC)`  
- Marks: numbers, `AB`, `OD` (on duty → treated as absent), `NA`

### Analysis

- Pass / fail / absentee totals (configurable pass mark, default **50**)
- Subject-wise: appeared, absent, passed, failed, pass %, average, highest
- Arrear groups: all clear, 1 subject, 2 subjects, 3+
- Category particulars: Counselling / Management × Hosteller / Day Scholar × Boys / Girls
- Ranked toppers and full student table with search & filters

### Diagrams (Overview)

- Subject-wise pass / fail bar chart
- Arrear distribution pie chart
- Pass vs fail overview chart
- Subject summary with progress bars

### Sidebar

- Editable assessment meta (title, department, year, semester, batch, section)
- Full subject names + staff + pass %
- Ranked student names with P/F

### Export Excel

- **FRONT** — institution header, particulars table, topper list, subject-wise performance, signature block
- **MARK** — full student list, summary rows, arrear section groups
- Professional borders, merged headers, AB (yellow), fail marks (red)
- Landscape, fit-to-page print layout

---

## Exam cell workflow

1. Open the portal (empty until you upload or load demo)
2. **Upload mark sheet** or **Load demo data**
3. Check the notice: student count and subject codes
4. Review **Overview** diagrams and **Subjects** table
5. (Optional) **Merge name list** for official spellings
6. Adjust pass mark in the sidebar if needed
7. **Export analysis** → submit the downloaded file

---

## Local development

```bash
npm install
npm run dev
```

```bash
npm run build
npm run preview
```

Stack: TanStack Start · React 19 · Tailwind 4 · ExcelJS · SheetJS · Recharts · Vercel

---

## Expected mark sheet columns

**Minimum**

| Column | Example |
|--------|---------|
| Reg.No | `310824148001` |
| Name | `STUDENT NAME` |
| Subject codes | `CS3452(TOC)` or `CS3452 & Theory of Computation` |

**Optional (for category tables)**

| Column | Values |
|--------|--------|
| B/G | B / G |
| C/M | C / M |
| H/DS | H / DS |

---

## Project structure (core)

```
src/
  lib/
    parse-sheet.ts   # Excel → students / subjects / meta
    analysis.ts      # Pass/fail, arrears, particulars, ranks
    export-sheet.ts  # FRONT + MARK professional export
  data/seed.ts       # Demo students & subjects
  routes/index.tsx   # Main portal UI
  components/
    result-charts.tsx
```

---

## License

Private academic / exam-cell use.
