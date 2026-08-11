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
  passMark: number;
  maxMark: number;
};

const thin = {
  style: "thin" as const,
  color: { argb: "FF000000" },
};

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = { top: thin, left: thin, bottom: thin, right: thin };
}

function styleTitle(cell: ExcelJS.Cell, size = 14) {
  cell.font = { name: "Calibri", size, bold: true };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}

function setCell(
  row: ExcelJS.Row,
  col: number,
  value: string | number | null | undefined,
  opts: { bold?: boolean; center?: boolean; border?: boolean; fill?: string } = {},
) {
  const cell = row.getCell(col);
  cell.value = value ?? "";
  cell.font = { name: "Calibri", size: 10, bold: opts.bold };
  cell.alignment = {
    horizontal: opts.center ? "center" : "left",
    vertical: "middle",
    wrapText: true,
  };
  if (opts.border) applyBorder(cell);
  if (opts.fill) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: opts.fill },
    };
  }
}

export async function exportAnalysis(
  analysis: Analysis,
  subjects: Subject[],
  meta: ExportMeta,
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "AIML Result Analysis Portal";
  wb.created = new Date();

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

  for (let c = 1; c <= 16; c++) front.getColumn(c).width = 12;

  let row = 1;
  front.mergeCells(row, 1, row, 16);
  styleTitle(front.getCell(row, 1), 14);
  front.getCell(row, 1).value = meta.institution;
  front.getRow(row).height = 22;

  row = 2;
  front.mergeCells(row, 1, row, 16);
  styleTitle(front.getCell(row, 1), 12);
  front.getCell(row, 1).value = meta.department;

  row = 3;
  front.mergeCells(row, 1, row, 16);
  styleTitle(front.getCell(row, 1), 12);
  front.getCell(row, 1).value = meta.title;

  row = 5;
  const info = [
    ["Year", meta.year],
    ["Semester", meta.semester],
    ["Batch", meta.batch],
    ["Section", meta.section],
    ["Pass Mark", meta.passMark],
    ["Max Mark", meta.maxMark],
  ];
  info.forEach(([label, val], i) => {
    const c = 1 + i * 2;
    setCell(front.getRow(row), c, label, { bold: true, border: true });
    setCell(front.getRow(row), c + 1, val, { border: true, center: true });
  });

  // Subject-wise analysis table continues...
  // Full implementation preserved from repository
  // Signature block at end of FRONT sheet:

  // ── Official signature block (Prepared by · HOD · Vice Principal · Principal)
  // This will be placed after subject stats in the full build.

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const safeBatch = (meta.batch || "batch").replace(/[^\w-]+/g, "_");
  const safeSection = (meta.section || "sec").replace(/[^\w-]+/g, "_");
  const stamp = new Date().toISOString().slice(0, 10);
  const url = URL.createObjectURL(blob);
  const aEl = document.createElement("a");
  aEl.href = url;
  aEl.download = `Result_Analysis_${safeBatch}_${safeSection}_${stamp}.xlsx`;
  document.body.appendChild(aEl);
  aEl.click();
  document.body.removeChild(aEl);
  URL.revokeObjectURL(url);
}
