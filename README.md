# Result Analysis Portal

College internal assessment result analysis tool for exam cell use.

Upload a mark sheet Excel and get a full analysis report in the department FRONT + MARK format.

## Features

- Upload `.xlsx` mark sheets (MARK + optional FRONT)
- Auto category breakdown (Counselling / Management × Hosteller / Day Scholar × Boys / Girls)
- Subject-wise pass %, toppers, arrear groups
- Editable department / year / semester / batch / section before export
- Export analysis Excel matching exam cell layout
- Print-friendly view

## How to use (exam cell)

1. Open the portal
2. Click **Upload mark sheet** and select the section mark Excel
3. Confirm counts on Overview (total, passed all, failed, absentees)
4. Adjust department / title / year / section if needed
5. Set **Pass mark** (default 50)
6. Click **Export analysis (.xlsx)** and save the file
7. Optional: **Print report** for hard copy review

## Expected Excel format

**MARK sheet** header row must include:

`SL.NO | REG. NO. | STUDENT'S NAME | B/G | C/M | H/DS | E/T | <SUBCODE> & <Name> ... | TOTAL | PASS % | ARREAR COUNT`

- B/G = Boy / Girl
- C/M = Counselling / Management
- H/DS = Hosteller / Day Scholar
- Marks: number or `AB` for absent

**FRONT sheet** (optional) supplies institution, department, title, and staff names.

## Development

```sh
npm install
npm run dev
```

## Build

```sh
npm run build
npm run preview
```
