import ExcelJS from "exceljs";
import { fmt, type Analysis, type Subject } from "@/lib/analysis";

export type ExportMeta = {
  institution: string;
  department: string;
  title: string;
  year: string;
  semester: string;
  batch: string;
  section: string;
};

const thin: Partial<ExcelJS.Border> = {
  style: "thin",
  color: { argb: "FF000000" },
};
const medium: Partial<ExcelJS.Border> = {
  style: "medium",
  color: { argb: "FF000000" },
};
const allThin: Partial<ExcelJS.Borders> = {
  top: thin,
  left: thin,
  bottom: thin,
  right: thin,
};
const allMedium: Partial<ExcelJS.Borders> = {
  top: medium,
  left: medium,
  bottom: medium,
  right: medium,
};

function applyBorder(
  cell: ExcelJS.Cell,
  border: Partial<ExcelJS.Borders> = allThin,
) {
  cell.border = border;
}

function styleTitle(cell: ExcelJS.Cell, size = 14) {
  cell.font = { bold: true, size, name: "Calibri" };
  cell.alignment = { horizontal: "center", vertical: "middle" };
}

function styleLeft(cell: ExcelJS.Cell) {
  cell.alignment = { horizontal: "left", vertical: "middle" };
  cell.font = { name: "Calibri", size: 10 };
}

function setCell(
  row: ExcelJS.Row,
  col: number,
  value: string | number | null | undefined,
  opts?: {
    border?: boolean;
    bold?: boolean;
    center?: boolean;
    header?: boolean;
  },
) {
  const cell = row.getCell(col);
  if (value !== null && value !== undefined && value !== "") {
    cell.value = value;
  }
  cell.font = {
    name: "Calibri",
    size: opts?.header ? 11 : 10,
    bold: opts?.bold || opts?.header || false,
  };
  cell.alignment = {
    horizontal: opts?.center || opts?.header ? "center" : "left",
    vertical: "middle",
    wrapText: true,
  };
  if (opts?.border || opts?.header) {
    applyBorder(cell, opts?.header ? allMedium : allThin);
  }
  if (opts?.header) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE8EEF7" },
    };
  }
  return cell;
}

export async function exportAnalysisSheet(
  a: Analysis,
  subjects: Subject[],
  meta: ExportMeta,
  passMark = 50,
) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Result Analysis Portal";
  wb.created = new Date();

  const cat = a.categoryCounts;
  const hostC = (cat["C-H-B"] ?? 0) + (cat["C-H-G"] ?? 0);
  const dayC = (cat["C-DS-B"] ?? 0) + (cat["C-DS-G"] ?? 0);
  const hostM = (cat["M-H-B"] ?? 0) + (cat["M-H-G"] ?? 0);
  const dayM = (cat["M-DS-B"] ?? 0) + (cat["M-DS-G"] ?? 0);

  const front = wb.addWorksheet("FRONT", {
    views: [{ showGridLines: false }],
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
    },
  });

  const frontWidths = [6, 42, 4, 4, 4, 4, 4, 10, 9, 9, 9, 9, 9, 9, 9, 9];
  frontWidths.forEach((w, i) => {
    front.getColumn(i + 1).width = w;
  });

  front.mergeCells(1, 1, 1, 16);
  styleTitle(front.getCell(1, 1), 13);
  front.getCell(1, 1).value = meta.institution;
  front.getRow(1).height = 22;

  front.mergeCells(2, 1, 2, 16);
  styleTitle(front.getCell(2, 1), 12);
  front.getCell(2, 1).value = meta.department;
  front.getRow(2).height = 20;

  front.mergeCells(3, 1, 3, 16);
  styleTitle(front.getCell(3, 1), 12);
  front.getCell(3, 1).value = meta.title;
  front.getRow(3).height = 20;

  const metaRow = front.getRow(4);
  setCell(metaRow, 1, "YEAR  :", { bold: true });
  setCell(metaRow, 2, meta.year, { bold: true, center: true });
  setCell(metaRow, 3, "SEMESTER:", { bold: true });
  setCell(metaRow, 4, meta.semester, { bold: true, center: true });
  setCell(metaRow, 5, "BATCH:", { bold: true });
  setCell(metaRow, 7, "BATCH:", { bold: true });
  setCell(metaRow, 8, meta.batch, { bold: true, center: true });
  setCell(metaRow, 11, "SECTION:", { bold: true });
  setCell(metaRow, 14, meta.section, { bold: true, center: true });
  metaRow.height = 18;

  front.getRow(5).height = 8;

  const h1 = front.getRow(6);
  h1.height = 18;
  setCell(h1, 1, "S.No.", { header: true });
  front.mergeCells(6, 2, 6, 7);
  setCell(h1, 2, "PARTICULARS", { header: true });
  setCell(h1, 8, "TOTAL", { header: true });
  front.mergeCells(6, 9, 6, 12);
  setCell(h1, 9, "COUNSELLING", { header: true });
  for (let c = 10; c <= 12; c++) applyBorder(h1.getCell(c), allMedium);
  front.mergeCells(6, 13, 6, 16);
  setCell(h1, 13, "MANAGEMENT", { header: true });
  for (let c = 14; c <= 16; c++) applyBorder(h1.getCell(c), allMedium);

  const h2 = front.getRow(7);
  h2.height = 18;
  for (let c = 1; c <= 8; c++) applyBorder(h2.getCell(c));
  front.mergeCells(7, 9, 7, 10);
  setCell(h2, 9, `HOSTELLER(${hostC})`, { header: true });
  applyBorder(h2.getCell(10), allMedium);
  front.mergeCells(7, 11, 7, 12);
  setCell(h2, 11, `DAY SCHOLAR(${dayC})`, { header: true });
  applyBorder(h2.getCell(12), allMedium);
  front.mergeCells(7, 13, 7, 14);
  setCell(h2, 13, `HOSTELLER(${hostM})`, { header: true });
  applyBorder(h2.getCell(14), allMedium);
  front.mergeCells(7, 15, 7, 16);
  setCell(h2, 15, `DAY SCHOLAR(${dayM})`, { header: true });
  applyBorder(h2.getCell(16), allMedium);

  const h3 = front.getRow(8);
  h3.height = 16;
  for (let c = 1; c <= 8; c++) applyBorder(h3.getCell(c));
  const genderLabels = ["BOYS", "GIRLS", "BOYS", "GIRLS", "BOYS", "GIRLS", "BOYS", "GIRLS"];
  genderLabels.forEach((g, i) => {
    setCell(h3, 9 + i, g, { header: true });
  });

  const catOrder = [
    "C-H-B",
    "C-H-G",
    "C-DS-B",
    "C-DS-G",
    "M-H-B",
    "M-H-G",
    "M-DS-B",
    "M-DS-G",
  ] as const;

  a.particulars.forEach((p, i) => {
    const r = front.getRow(9 + i);
    r.height = 17;
    setCell(r, 1, i + 1, { border: true, center: true });
    front.mergeCells(9 + i, 2, 9 + i, 7);
    setCell(r, 2, p.label, { border: true });
    for (let c = 3; c <= 7; c++) applyBorder(r.getCell(c));
    setCell(r, 8, p.total, { border: true, center: true, bold: true });
    catOrder.forEach((k, ci) => {
      setCell(r, 9 + ci, p.byCategory[k], { border: true, center: true });
    });
  });

  const passRowIdx = 9 + a.particulars.length;
  const passRow = front.getRow(passRowIdx);
  passRow.height = 17;
  setCell(passRow, 1, a.particulars.length + 1, { border: true, center: true });
  front.mergeCells(passRowIdx, 2, passRowIdx, 7);
  setCell(passRow, 2, "PASS PERCENTAGE", { border: true, bold: true });
  for (let c = 3; c <= 7; c++) applyBorder(passRow.getCell(c));
  setCell(passRow, 8, Number(fmt(a.passPercent, 2)), {
    border: true,
    center: true,
    bold: true,
  });
  for (let c = 9; c <= 16; c++) applyBorder(passRow.getCell(c));

  let row = passRowIdx + 2;
  front.mergeCells(row, 1, row, 16);
  styleTitle(front.getCell(row, 1), 12);
  front.getCell(row, 1).value = "TOPPER LIST";
  front.getRow(row).height = 20;

  row += 1;
  const topH = front.getRow(row);
  topH.height = 17;
  for (let c = 1; c <= 16; c++) applyBorder(topH.getCell(c), allMedium);
  setCell(topH, 2, "S.NO", { header: true });
  setCell(topH, 3, "REG.NO", { header: true });
  setCell(topH, 4, "PHOTO", { header: true });
  front.mergeCells(row, 5, row, 9);
  setCell(topH, 5, "STUDENTS NAME", { header: true });
  setCell(topH, 10, "TOTAL", { header: true });
  setCell(topH, 12, "PASS %", { header: true });
  setCell(topH, 14, "RANK", { header: true });

  a.toppers.slice(0, 3).forEach((t, i) => {
    row += 1;
    const tr = front.getRow(row);
    tr.height = 18;
    const rankLabel =
      t.rank === 1 ? "I" : t.rank === 2 ? "II" : t.rank === 3 ? "III" : String(t.rank);
    setCell(tr, 2, i + 1, { border: true, center: true });
    setCell(tr, 3, t.student.reg, { border: true, center: true });
    applyBorder(tr.getCell(4));
    front.mergeCells(row, 5, row, 9);
    setCell(tr, 5, t.student.name, { border: true });
    for (let c = 6; c <= 9; c++) applyBorder(tr.getCell(c));
    setCell(tr, 10, t.total, { border: true, center: true, bold: true });
    applyBorder(tr.getCell(11));
    setCell(tr, 12, Number(fmt(t.percent, 2)), { border: true, center: true });
    applyBorder(tr.getCell(13));
    setCell(tr, 14, rankLabel, { border: true, center: true, bold: true });
  });

  row += 2;
  front.mergeCells(row, 1, row, 16);
  styleTitle(front.getCell(row, 1), 12);
  front.getCell(row, 1).value = "SUBJECT WISE PERFORMANCE";
  front.getRow(row).height = 20;

  row += 1;
  const subH = front.getRow(row);
  subH.height = 18;
  for (let c = 1; c <= 16; c++) applyBorder(subH.getCell(c), allMedium);
  setCell(subH, 2, "S.NO", { header: true });
  setCell(subH, 3, "SUB CODE", { header: true });
  front.mergeCells(row, 4, row, 5);
  setCell(subH, 4, "SUBJECT NAME", { header: true });
  front.mergeCells(row, 6, row, 9);
  setCell(subH, 6, "STAFF NAME", { header: true });
  setCell(subH, 10, "APPEARED", { header: true });
  setCell(subH, 11, "ABSENT", { header: true });
  setCell(subH, 12, "PASSED", { header: true });
  setCell(subH, 13, "FAILED", { header: true });
  setCell(subH, 14, "PASS %", { header: true });

  a.subjectStats.forEach((s, i) => {
    row += 1;
    const sr = front.getRow(row);
    sr.height = 17;
    for (let c = 1; c <= 16; c++) applyBorder(sr.getCell(c));
    setCell(sr, 2, i + 1, { border: true, center: true });
    setCell(sr, 3, s.subject.code, { border: true, center: true });
    front.mergeCells(row, 4, row, 5);
    setCell(sr, 4, s.subject.name, { border: true });
    front.mergeCells(row, 6, row, 9);
    setCell(sr, 6, s.subject.staff || "", { border: true });
    setCell(sr, 10, s.appeared + s.absent, { border: true, center: true });
    setCell(sr, 11, s.absent, { border: true, center: true });
    setCell(sr, 12, s.passed, { border: true, center: true });
    setCell(sr, 13, s.failed, { border: true, center: true });
    setCell(sr, 14, Math.round(s.passPercent), { border: true, center: true });
  });

  row += 4;
  const sig = front.getRow(row);
  sig.height = 22;
  setCell(sig, 2, "PREPARED BY", { bold: true, center: true });
  setCell(sig, 5, "HOD", { bold: true, center: true });
  setCell(sig, 10, "VICE PRINCIPAL", { bold: true, center: true });
  setCell(sig, 14, "PRINCIPAL", { bold: true, center: true });

  const mark = wb.addWorksheet("MARK", {
    views: [{ showGridLines: false }],
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
    },
  });

  const subjectHeaders = subjects.map((s) => `${s.code} & ${s.name}`);
  const markCols = 7 + subjects.length + 3;

  mark.getColumn(1).width = 6;
  mark.getColumn(2).width = 14;
  mark.getColumn(3).width = 22;
  mark.getColumn(4).width = 5;
  mark.getColumn(5).width = 5;
  mark.getColumn(6).width = 6;
  mark.getColumn(7).width = 5;
  for (let i = 0; i < subjects.length; i++) {
    mark.getColumn(8 + i).width = 12;
  }
  mark.getColumn(8 + subjects.length).width = 8;
  mark.getColumn(9 + subjects.length).width = 9;
  mark.getColumn(10 + subjects.length).width = 10;

  const writeMarkHeaderBlock = (startRow: number, blockTitle?: string) => {
    let r = startRow;
    if (blockTitle) {
      mark.mergeCells(r, 1, r, markCols);
      styleTitle(mark.getCell(r, 1), 12);
      mark.getCell(r, 1).value = blockTitle;
      mark.getRow(r).height = 18;
      r += 1;
    }

    const headers = [
      "SL.NO",
      "REG. NO.",
      "STUDENT'S NAME",
      "B/G",
      "C/M",
      "H/DS",
      "E/T",
      ...subjectHeaders,
      "TOTAL",
      "PASS %",
      "ARREAR COUNT",
    ];
    const hr = mark.getRow(r);
    hr.height = 30;
    headers.forEach((h, i) => {
      const cell = setCell(hr, i + 1, h, { header: true });
      cell.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
    });
    return r;
  };

  const writeStudentRows = (startRow: number, list: typeof a.results): number => {
    let r = startRow;
    list.forEach((res, i) => {
      const row = mark.getRow(r);
      row.height = 16;
      const vals: (string | number)[] = [
        i + 1,
        res.student.reg,
        res.student.name,
        res.student.gender,
        res.student.quota,
        res.student.stay,
        "E",
        ...res.student.marks.map((m) => (m === null ? "" : m)),
        res.total,
        Number(fmt(res.percent, 2)),
        res.arrears,
      ];
      vals.forEach((v, ci) => {
        const cell = setCell(row, ci + 1, v, {
          border: true,
          center: ci !== 2,
        });
        if (ci === 2) styleLeft(cell);
        if (ci >= 7 && ci < 7 + subjects.length) {
          const m = res.student.marks[ci - 7];
          if (m === "AB") {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFFF3CD" },
            };
          } else if (typeof m === "number" && m < passMark) {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFF8D7DA" },
            };
          }
        }
      });
      r += 1;
    });
    return r;
  };

  mark.mergeCells(1, 1, 1, markCols);
  styleTitle(mark.getCell(1, 1), 13);
  mark.getCell(1, 1).value = meta.institution;

  mark.mergeCells(2, 1, 2, markCols);
  styleTitle(mark.getCell(2, 1), 12);
  mark.getCell(2, 1).value = meta.department;

  mark.mergeCells(3, 1, 3, markCols);
  styleTitle(mark.getCell(3, 1), 12);
  mark.getCell(3, 1).value = "ASSESSMENT RESULT ANALYSIS";

  const mr = mark.getRow(4);
  setCell(mr, 1, "YEAR  :", { bold: true });
  setCell(mr, 2, meta.year, { bold: true, center: true });
  setCell(mr, 3, "SEMESTER:", { bold: true });
  setCell(mr, 4, meta.semester, { bold: true, center: true });
  setCell(mr, 5, "BATCH:", { bold: true });
  setCell(mr, 7, "BATCH:", { bold: true });
  setCell(mr, 8, meta.batch, { bold: true, center: true });
  setCell(mr, 10, "SECTION:", { bold: true });
  setCell(mr, 12, meta.section, { bold: true, center: true });

  let r = writeMarkHeaderBlock(6);
  const sorted = [...a.results].sort((x, y) =>
    x.student.reg.localeCompare(y.student.reg),
  );
  r = writeStudentRows(r + 1, sorted);

  const summaryLabels = [
    ["Total No. of Students", subjects.map(() => a.totalStudents)],
    ["No. of Students Present", a.subjectStats.map((s) => s.appeared)],
    ["No. of Students Absent", a.subjectStats.map((s) => s.absent)],
    ["No. of Students Pass", a.subjectStats.map((s) => s.passed)],
    ["No. of Students fail", a.subjectStats.map((s) => s.failed)],
    ["PASS%", a.subjectStats.map((s) => Number(fmt(s.passPercent, 2)))],
  ] as const;

  summaryLabels.forEach(([label, vals]) => {
    const row = mark.getRow(r);
    row.height = 16;
    setCell(row, 1, label, { border: true, bold: true });
    for (let c = 2; c <= 7; c++) applyBorder(row.getCell(c));
    vals.forEach((v, i) => {
      setCell(row, 8 + i, v, { border: true, center: true, bold: true });
    });
    for (let c = 8 + subjects.length; c <= markCols; c++) {
      applyBorder(row.getCell(c));
    }
    r += 1;
  });

  const groups: [string, typeof a.results][] = [
    ["ALL CLEAR STUDENTS", sorted.filter((x) => x.arrears === 0)],
    ["ONE SUBJECTS ARREAR STUDENTS", sorted.filter((x) => x.arrears === 1)],
    ["TWO SUBJECTS ARREAR STUDENTS", sorted.filter((x) => x.arrears === 2)],
    [
      "THREE & ABOVE SUBJECTS ARREAR STUDENTS",
      sorted.filter((x) => x.arrears >= 3),
    ],
  ];

  groups.forEach(([title, list]) => {
    r += 1;
    mark.mergeCells(r, 1, r, markCols);
    styleTitle(mark.getCell(r, 1), 12);
    mark.getCell(r, 1).value = meta.institution;
    r += 1;
    mark.mergeCells(r, 1, r, markCols);
    styleTitle(mark.getCell(r, 1), 11);
    mark.getCell(r, 1).value = meta.department;
    r += 1;
    mark.mergeCells(r, 1, r, markCols);
    styleTitle(mark.getCell(r, 1), 11);
    mark.getCell(r, 1).value = "ASSESSMENT RESULT ANALYSIS";
    r += 1;
    r = writeMarkHeaderBlock(r, title);
    r = writeStudentRows(r + 1, list);
  });

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const safeBatch = (meta.batch || "batch").replace(/[^\w-]+/g, "_");
  const safeSection = (meta.section || "sec").replace(/[^\w-]+/g, "_");
  const url = URL.createObjectURL(blob);
  const aEl = document.createElement("a");
  aEl.href = url;
  aEl.download = `Result_Analysis_${safeBatch}_${safeSection}.xlsx`;
  document.body.appendChild(aEl);
  aEl.click();
  document.body.removeChild(aEl);
  URL.revokeObjectURL(url);
}
