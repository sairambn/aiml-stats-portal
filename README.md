<div align="center">

# Result Analysis Portal

**Exam-cell ready** internal assessment result analysis.

Upload a section mark sheet → view diagrams and tables → download a professional **FRONT + MARK** Excel report.

[![Live](https://img.shields.io/badge/Live-aiml--stats--portal.vercel.app-black?style=for-the-badge&logo=vercel)](https://aiml-stats-portal.vercel.app)
[![React 19](https://img.shields.io/badge/React-19-61DAFB?style=flat-square&logo=react&logoColor=black)](https://react.dev)
[![TypeScript](https://img.shields.io/badge/TypeScript-5-3178C6?style=flat-square&logo=typescript&logoColor=white)](https://www.typescriptlang.org)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](./LICENSE)

**Live:** [aiml-stats-portal.vercel.app](https://aiml-stats-portal.vercel.app)

</div>

---

## Privacy first

No student data is stored on the server. Everything runs in the browser session. Refresh clears the session.

---

## Quick start (30 seconds)

1. Open the [live portal](https://aiml-stats-portal.vercel.app) **or** run locally.
2. Upload a mark sheet (`.xlsx`) — or click **Load demo data** to try instantly.
3. Review Overview diagrams, Subjects, Students, Toppers.
4. Adjust pass mark in the sidebar if needed (default **50**).
5. **Export analysis** → exam-cell Excel with borders, AB (yellow), fails (red).

---

## Sample files for testing

| File | Contents |
|------|----------|
| `RA II AIML.xlsx` / `RA III AIML.xlsx` | Your section mark sheets with date + short-code headers (DM, OS, NLP, …) |

Use the built-in **Load demo data** button or upload the files you already have.

---

## What it does

| Step | Action |
|------|--------|
| 1 | **Upload mark sheet** (`.xlsx`) |
| 2 | Portal reads students, subjects, marks, absents |
| 3 | Overview **diagrams** + pass / fail / arrear stats |
| 4 | **Export analysis** → exam-cell Excel with borders |

---

## Features

### Upload formats supported

- **Full RA sheet** — `FRONT` + `MARK` with B/G, C/M, H/DS, category tables  
- **Simple IAT / AIML mark sheet** — `S.No | Reg.No | Name | CS3452(TOC) | …`  
- **Date + short-code header** (common AIML format)  
  ```
  S.No | Register No. | Name | 30/07/26 | 31/07/26 | …
                 | DM       | OS       | …
  ```
  Codes like `DM`, `OS`, `OOSE`, `DS`, `JP`, `NLP`, `DL`, `DC`, `CCS`, `CC`, `STA` are recognised and expanded to full names.
- Subject headers: `CS3452 & Theory of Computation` **or** `CS3452(TOC)` **or** short codes under dates  
- Marks: numbers, `AB` / `Ab`, `OD` (on duty → treated as absent), `NA`

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

### Sidebar

- Editable assessment meta (title, department, year, semester, batch, section)
- Full subject names
- Configurable pass mark

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
5. Adjust pass mark in the sidebar if needed
6. **Export analysis** → submit the downloaded file

---

## Local development

```bash
bun install          # or: npm install
bun run dev          # or: npm run dev
```

```bash
bun run build && bun run preview
```

Stack: **TanStack Start** · React 19 · Tailwind 4 · ExcelJS · SheetJS · Recharts · Vercel

---

## Expected mark sheet columns

**Minimum**

| Column | Example |
|--------|---------|
| Reg.No | `310824148001` or `25JELAIML402` |
| Name | `STUDENT NAME` |
| Subject codes | `CS3452(TOC)` **or** short codes (`DM`, `NLP`…) under date headers |

**Optional (for category tables)**

| Column | Values |
|--------|--------|
| B/G | B / G |
| C/M | C / M |
| H/DS | H / DS |

---

## Project structure (core)

```text
src/
  lib/
    parse-sheet.ts   # Excel → students / subjects / meta
    analysis.ts      # Pass/fail, arrears, particulars, ranks
    export-sheet.ts  # FRONT + MARK professional export
  data/seed.ts       # Demo students & subjects
  routes/
    portal.tsx       # Main portal UI
  components/
    result-charts.tsx
```

---

## License

[MIT](./LICENSE) — free for academic and exam-cell use.

Built by [@sairambn](https://github.com/sairambn) · **SR**
