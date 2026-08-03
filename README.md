# Result Analysis Portal

Professional internal assessment result analysis for college exam cell use.

Upload a section mark sheet and generate the official **FRONT + MARK** analysis workbook with category breakdown, subject-wise performance, toppers and arrear groups.

---

## Features

| Area | What you get |
|------|----------------|
| **Upload** | Reads MARK (and FRONT) sheets — students, B/G, C/M, H/DS, marks, AB |
| **Sidebar** | Full subject names + staff, ranked student names, batch meta, pass mark |
| **Overview** | Pass %, arrear distribution, class average |
| **Particulars** | Counselling / Management × Hosteller / Day Scholar × Boys / Girls |
| **Subjects** | Appeared, absent, passed, failed, pass %, average, highest |
| **Students** | Rank, marks, arrears, search & filter |
| **Toppers** | Top ranks with totals |
| **Export** | Excel with borders, merged headers, highlights — exam cell ready |

---

## Quick start (exam cell)

1. Open the deployed portal  
2. **Upload mark sheet** → select the section `.xlsx`  
3. Confirm totals on Overview  
4. Adjust pass mark if needed (default **50**)  
5. **Export analysis** → download FRONT + MARK workbook  
6. Submit the file to exam cell  

---

## Expected mark sheet format

Header row on **MARK**:

```
SL.NO | REG. NO. | STUDENT'S NAME | B/G | C/M | H/DS | E/T | <CODE> & <Subject Name> … | TOTAL | PASS % | ARREAR COUNT
```

| Column | Meaning |
|--------|---------|
| B/G | Boy / Girl |
| C/M | Counselling / Management |
| H/DS | Hosteller / Day Scholar |
| Marks | Number, or `AB` for absent |

Optional **FRONT** sheet supplies institution, department, title, year, semester, batch, section and staff names.

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

---

## Tech

- TanStack Start (React + Vite)  
- ExcelJS (styled export with borders)  
- SheetJS / xlsx (parse)  
- Tailwind CSS  

---

## License

Private institutional use — Result Analysis Portal.
