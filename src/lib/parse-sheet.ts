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

  const r0 = cellStr(grid[0]?.[0]);
  if (r0) meta.institution = r0;

  const r1 = cellStr(grid[1]?.[0]);
  if (r1 && /department/i.test(r1)) meta.department = r1;

  const r2 = cellStr(grid[2]?.[0]);
  if (r2 && /assessment|result|analysis/i.test(r2)) meta.title = r2;

  // Scan first 8 rows for YEAR / SEMESTER / BATCH / SECTION
  for (let i = 0; i < Math.min(8, grid.length); i++) {
    const row = grid[i] as unknown[];
    for (let c = 0; c < (row?.length ?? 0); c++) {
      const label = cellStr(row[c]).toUpperCase().replace(/\s+/g, " ");
      const next = cellStr(row[c + 1]);
      if (label.startsWith("YEAR") && next) meta.year = next;
      if (label.startsWith("SEMESTER") && next) meta.semester = next;
      if (label.startsWith("BATCH") && next && /\d{4}/.test(next)) meta.batch = next;
      if (label.startsWith("SECTION") && next) meta.section = next;
    }
  }
  return meta;
}

function parseStaffFromFront(grid: unknown[][]): Record<string, string> {
  const staffMap: Record<string, string> = {};
  const subjIdx = grid.findIndex((row) =>
    row.some((c) => typeof c === "string" && /subject\s*wise/i.test(c)),
  );
  if (subjIdx < 0) return staffMap;

  for (const row of grid.slice(subjIdx + 1, subjIdx + 20)) {
    const cells = row as unknown[];
    // Find a subject code like CS3452 in the row
    for (let c = 0; c < cells.length; c++) {
      const code = cellStr(cells[c]).match(/^([A-Z]{2}\d{4})$/);
      if (!code) continue;
      // Staff is usually a few columns after subject name
      for (let k = c + 1; k < Math.min(c + 6, cells.length); k++) {
        const s = cellStr(cells[k]);
        if (/\b(AP|PROF|DR\.|MR\.|MS\.|MRS\.)/i.test(s) || /\//.test(s)) {
          staffMap[code[1]] = s;
          break;
        }
      }
    }
  }
  return staffMap;
}

export function parseWorkbook(data: ArrayBuffer): ParsedSheet {
  const wb = XLSX.read(data, { type: "array" });

  const markSheet =
    wb.Sheets["MARK"] ??
    wb.Sheets[wb.SheetNames.find((n) => /mark/i.test(n)) ?? ""] ??
    wb.Sheets[wb.SheetNames[wb.SheetNames.length - 1] ?? ""];

  if (!markSheet) throw new Error("No worksheet found in the file");

  const markGrid = XLSX.utils.sheet_to_json<unknown[]>(markSheet, {
    header: 1,
    defval: "",
  });

  const meta = parseMetaFromGrid(markGrid);

  // FRONT sheet for department title and staff names
  const frontSheet = wb.Sheets["FRONT"] ?? wb.Sheets[wb.SheetNames[0] ?? ""];
  let staffMap: Record<string, string> = {};
  if (frontSheet) {
    const frontGrid = XLSX.utils.sheet_to_json<unknown[]>(frontSheet, {
      header: 1,
      defval: "",
    });
    const frontMeta = parseMetaFromGrid(frontGrid);
    Object.assign(meta, frontMeta);
    staffMap = parseStaffFromFront(frontGrid);
  }

  // Find the first header row with REG. NO.
  const headerIdx = markGrid.findIndex((row) =>
    row.some((c) => typeof c === "string" && /reg\.?\s*no/i.test(String(c))),
  );
  if (headerIdx < 0) throw new Error("Could not find a REG. NO. header row");

  const header = markGrid[headerIdx] as unknown[];

  // Map column indices for identity fields
  let colReg = 1;
  let colName = 2;
  let colGender = 3;
  let colQuota = 4;
  let colStay = 5;

  header.forEach((cell, col) => {
    const h = cellStr(cell).toUpperCase();
    if (/REG/.test(h) && /NO/.test(h)) colReg = col;
    if (/NAME/.test(h) && !/SUB/.test(h)) colName = col;
    if (/^B\/?G$/.test(h) || h === "B/G") colGender = col;
    if (/^C\/?M$/.test(h) || h === "C/M" || h === "G/M") colQuota = col;
    if (/^H\/?DS$/.test(h) || h === "H/DS") colStay = col;
  });

  // Subject columns: code pattern CS3452 or CS3452 & Name or CS3451& Name
  const subjectCols: { col: number; subject: Subject }[] = [];
  header.forEach((cell, col) => {
    if (typeof cell !== "string") return;
    if (col <= colStay) return;
    const m = cell.match(/^\s*([A-Z]{2}\d{4})\s*&?\s*(.*)$/i);
    if (!m) return;
    const code = m[1].toUpperCase();
    const name = (m[2] ?? "").trim() || code;
    // Skip TOTAL / PASS / ARREAR columns
    if (/total|pass|arrear|absent|status/i.test(name) && !/theory|system|machine|web|operating|environ/i.test(name)) {
      return;
    }
    subjectCols.push({
      col,
      subject: {
        code,
        name,
        staff: staffMap[code] ?? "",
      },
    });
  });

  if (!subjectCols.length) throw new Error("No subject columns found in the mark sheet");

  const parsed: Student[] = [];
  const seen = new Set<string>();

  for (const row of markGrid.slice(headerIdx + 1)) {
    const cells = row as unknown[];
    const first = cellStr(cells[0]).toUpperCase();
    // Stop at summary / section headers of the mark sheet
    if (
      /total\s*no/i.test(first) ||
      /no\.?\s*of\s*students/i.test(first) ||
      /all\s*clear/i.test(first) ||
      /one\s*subject/i.test(first) ||
      /two\s*subject/i.test(first) ||
      /three/i.test(first) ||
      /arrear/i.test(first) ||
      /pass%/i.test(first) ||
      /jeppiaar/i.test(first) ||
      /department/i.test(first) ||
      /assessment/i.test(first)
    ) {
      break;
    }

    const reg = cellStr(cells[colReg]);
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
