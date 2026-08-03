import * as XLSX from "xlsx";
import type { Mark, Student, Subject } from "@/lib/analysis";
import type { ExportMeta } from "@/lib/export-sheet";

export type ParsedSheet = {
  students: Student[];
  subjects: Subject[];
  meta: Partial<ExportMeta>;
};

function cellStr(v: unknown): string {
  if (v === undefined || v === null) return "";
  return String(v).trim();
}

function parseMark(v: unknown): Mark {
  if (v === undefined || v === null || v === "") return null;
  if (typeof v === "number") return Math.round(v);
  const s = String(v).trim();
  if (/^ab$/i.test(s)) return "AB";
  const n = Number(s);
  return Number.isFinite(n) ? Math.round(n) : 0;
}

function parseMetaFromGrid(grid: unknown[][]): Partial<ExportMeta> {
  const meta: Partial<ExportMeta> = {};
  if (!grid.length) return meta;

  // Scan top rows by content (first row is often empty in these sheets)
  for (let i = 0; i < Math.min(10, grid.length); i++) {
    const row = (grid[i] ?? []) as unknown[];
    const first = cellStr(row[0]);

    if (!meta.institution && first && /nagar|chennai|college|institute|university|road|salai/i.test(first)) {
      meta.institution = first;
    }
    if (!meta.department && first && /department/i.test(first)) {
      meta.department = first;
    }
    if (
      !meta.title &&
      first &&
      /assessment|result\s*analysis|internal/i.test(first) &&
      !/department/i.test(first)
    ) {
      meta.title = first;
    }

    for (let c = 0; c < row.length; c++) {
      const label = cellStr(row[c]).toUpperCase().replace(/\s+/g, " ");
      const next = cellStr(row[c + 1]);
      if (!next) continue;
      if (label.startsWith("YEAR") && !meta.year) meta.year = next;
      if (label.startsWith("SEMESTER") && !meta.semester) meta.semester = next;
      if (label.startsWith("BATCH") && /\d{4}/.test(next)) meta.batch = next;
      if (label.startsWith("SECTION") && !meta.section) meta.section = next;
    }
  }
  return meta;
}

function parseStaffFromFront(grid: unknown[][]): Record<string, string> {
  const staffMap: Record<string, string> = {};
  const subjIdx = grid.findIndex((row) =>
    row.some((c) => typeof c === "string" && /subject\s*wise/i.test(String(c))),
  );
  if (subjIdx < 0) return staffMap;

  for (const row of grid.slice(subjIdx + 1, subjIdx + 25)) {
    const cells = row as unknown[];
    for (let c = 0; c < cells.length; c++) {
      const codeMatch = cellStr(cells[c]).match(/^([A-Z]{2}\d{4})$/i);
      if (!codeMatch) continue;
      const code = codeMatch[1].toUpperCase();
      for (let k = c + 1; k < Math.min(c + 8, cells.length); k++) {
        const s = cellStr(cells[k]);
        if (!s) continue;
        if (/\b(AP|PROF|DR\.?|MR\.?|MS\.?|MRS\.?)\b/i.test(s) || /\//.test(s)) {
          staffMap[code] = s;
          break;
        }
      }
    }
  }
  return staffMap;
}

export function parseWorkbook(data: ArrayBuffer): ParsedSheet {
  const wb = XLSX.read(data, { type: "array", cellDates: true });

  const markName =
    wb.SheetNames.find((n) => /^mark$/i.test(n)) ??
    wb.SheetNames.find((n) => /mark/i.test(n)) ??
    wb.SheetNames[wb.SheetNames.length - 1];

  const markSheet = markName ? wb.Sheets[markName] : undefined;
  if (!markSheet) throw new Error("No worksheet found in the file");

  const markGrid = XLSX.utils.sheet_to_json<unknown[]>(markSheet, {
    header: 1,
    defval: "",
    raw: true,
  });

  const meta = parseMetaFromGrid(markGrid);

  const frontName =
    wb.SheetNames.find((n) => /^front$/i.test(n)) ??
    wb.SheetNames.find((n) => /front/i.test(n)) ??
    wb.SheetNames[0];

  let staffMap: Record<string, string> = {};
  if (frontName && wb.Sheets[frontName]) {
    const frontGrid = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[frontName], {
      header: 1,
      defval: "",
      raw: true,
    });
    const frontMeta = parseMetaFromGrid(frontGrid);
    // Prefer FRONT for title/department (more complete analysis header)
    if (frontMeta.institution) meta.institution = frontMeta.institution;
    if (frontMeta.department) meta.department = frontMeta.department;
    if (frontMeta.title) meta.title = frontMeta.title;
    if (frontMeta.year) meta.year = frontMeta.year;
    if (frontMeta.semester) meta.semester = frontMeta.semester;
    if (frontMeta.batch) meta.batch = frontMeta.batch;
    if (frontMeta.section) meta.section = frontMeta.section;
    staffMap = parseStaffFromFront(frontGrid);
  }

  const headerIdx = markGrid.findIndex((row) =>
    row.some((c) => typeof c === "string" && /reg\.?\s*no/i.test(String(c))),
  );
  if (headerIdx < 0) throw new Error("Could not find a REG. NO. header row");

  const header = markGrid[headerIdx] as unknown[];

  let colReg = 1;
  let colName = 2;
  let colGender = 3;
  let colQuota = 4;
  let colStay = 5;

  header.forEach((cell, col) => {
    const h = cellStr(cell).toUpperCase();
    if (/REG/.test(h) && /NO/.test(h)) colReg = col;
    if (/NAME/.test(h) && !/SUB/.test(h)) colName = col;
    if (/^B\s*\/?\s*G$/.test(h) || h === "B/G") colGender = col;
    if (/^C\s*\/?\s*M$/.test(h) || h === "C/M" || h === "G/M") colQuota = col;
    if (/^H\s*\/?\s*DS$/.test(h) || h === "H/DS") colStay = col;
  });

  const subjectCols: { col: number; subject: Subject }[] = [];
  header.forEach((cell, col) => {
    if (col <= colStay) return;
    const text = cellStr(cell);
    if (!text) return;
    const m = text.match(/^\s*([A-Z]{2}\d{4})\s*&?\s*(.*)$/i);
    if (!m) return;
    const code = m[1].toUpperCase();
    let name = (m[2] ?? "").trim();
    // Guard: don't treat TOTAL / PASS % as subjects
    if (!name && /total|pass|arrear/i.test(text)) return;
    if (/^(total|pass\s*%?|arrear.*|absent|status)$/i.test(name)) return;
    if (!name) name = code;
    subjectCols.push({
      col,
      subject: {
        code,
        name,
        staff: staffMap[code] ?? "",
      },
    });
  });

  if (!subjectCols.length) {
    throw new Error("No subject columns found in the mark sheet");
  }

  const parsed: Student[] = [];
  const seen = new Set<string>();

  for (const row of markGrid.slice(headerIdx + 1)) {
    const cells = row as unknown[];
    const first = cellStr(cells[0]).toUpperCase();

    if (
      /total\s*no/i.test(first) ||
      /no\.?\s*of\s*students/i.test(first) ||
      /all\s*clear/i.test(first) ||
      /one\s*subject/i.test(first) ||
      /two\s*subject/i.test(first) ||
      /three/i.test(first) ||
      /^arrear/i.test(first) ||
      /^pass%/i.test(first) ||
      /jeppiaar/i.test(first) ||
      /department/i.test(first) ||
      /assessment/i.test(first)
    ) {
      break;
    }

    // Reg can be number or string from Excel
    const reg = cellStr(cells[colReg]).replace(/\.0$/, "");
    if (!/^\d{6,}$/.test(reg)) continue;
    if (seen.has(reg)) continue;
    seen.add(reg);

    const genderRaw = cellStr(cells[colGender]).toUpperCase();
    const quotaRaw = cellStr(cells[colQuota]).toUpperCase();
    const stayRaw = cellStr(cells[colStay]).toUpperCase();

    parsed.push({
      reg,
      name: cellStr(cells[colName]),
      gender: genderRaw.startsWith("G") ? "G" : "B",
      quota: quotaRaw.startsWith("M") ? "M" : "C",
      stay: stayRaw.startsWith("H") ? "H" : "DS",
      marks: subjectCols.map(({ col }) => parseMark(cells[col])),
    });
  }

  if (!parsed.length) throw new Error("No student rows found in the mark sheet");

  return {
    students: parsed,
    subjects: subjectCols.map((s) => s.subject),
    meta,
  };
}
