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

const CODE_RE = /([A-Z]{2,4}\d{3,4})/i;

function parseMetaFromGrid(grid: unknown[][]): Partial<ExportMeta> {
  const meta: Partial<ExportMeta> = {};
  if (!grid.length) return meta;

  for (let i = 0; i < Math.min(12, grid.length); i++) {
    const row = (grid[i] ?? []) as unknown[];
    const first = cellStr(row[0]);

    if (
      !meta.institution &&
      first &&
      /nagar|chennai|college|institute|university|road|salai/i.test(first)
    ) {
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

function parseSubjectsFromFront(
  grid: unknown[][],
): Map<string, { name: string; staff: string }> {
  const map = new Map<string, { name: string; staff: string }>();
  const subjIdx = grid.findIndex((row) =>
    row.some((c) => typeof c === "string" && /subject\s*wise/i.test(String(c))),
  );
  if (subjIdx < 0) return map;

  for (const row of grid.slice(subjIdx + 1, subjIdx + 40)) {
    const cells = row as unknown[];
    for (let c = 0; c < cells.length; c++) {
      const codeMatch = cellStr(cells[c]).match(/^([A-Z]{2,4}\d{3,4})$/i);
      if (!codeMatch) continue;
      const code = codeMatch[1].toUpperCase();
      let name = "";
      let staff = "";
      for (let k = c + 1; k < Math.min(c + 12, cells.length); k++) {
        const s = cellStr(cells[k]);
        if (!s) continue;
        if (
          !name &&
          !/^\d+$/.test(s) &&
          !/\b(AP|PROF|DR\.?|MR\.?|MS\.?|MRS\.?)\b/i.test(s)
        ) {
          name = s.replace(/^\s*&\s*/, "").trim();
          continue;
        }
        if (/\b(AP|PROF|DR\.?|MR\.?|MS\.?|MRS\.?)\b/i.test(s) || /\//.test(s)) {
          staff = s;
          break;
        }
      }
      if (!map.has(code)) {
        map.set(code, { name: name || code, staff });
      }
    }
  }
  return map;
}

function isIdentityHeader(h: string): boolean {
  const u = h.toUpperCase().replace(/\s+/g, " ");
  return (
    /^(SL\.?\s*NO\.?|S\.?NO\.?|SNO)$/.test(u) ||
    (/REG/.test(u) && /NO/.test(u)) ||
    (/NAME/.test(u) && !/SUB/.test(u) && !CODE_RE.test(u)) ||
    /^(B\/?G|C\/?M|G\/?M|H\/?DS|E\/?T)$/.test(u) ||
    /^(TOTAL|PASS\s*%?|ARREAR.*|ABSENT|STATUS)$/.test(u)
  );
}

function extractSubjectFromHeader(
  text: string,
): { code: string; name: string } | null {
  if (!text || isIdentityHeader(text)) return null;

  // Primary: starts with subject code
  const m = text.match(/^\s*([A-Z]{2,4}\d{3,4})\s*[&\-–:]*\s*(.*)$/i);
  if (m) {
    const code = m[1].toUpperCase();
    let name = (m[2] ?? "").trim().replace(/^[&\-–:]+\s*/, "");
    if (/^(total|pass\s*%?|arrear.*|absent|status)$/i.test(name)) return null;
    return { code, name: name || code };
  }

  // Fallback: code anywhere in cell
  const anywhere = text.match(/\b([A-Z]{2,4}\d{3,4})\b/i);
  if (!anywhere) return null;
  const code = anywhere[1].toUpperCase();
  const name = text
    .replace(anywhere[0], "")
    .replace(/^[&\-–:\s]+/, "")
    .trim();
  if (/^(total|pass|arrear|absent|status)$/i.test(name)) return null;
  return { code, name: name || code };
}

function collectSubjectsFromRow(
  header: unknown[],
  skipCols: Set<number>,
  frontSubjects: Map<string, { name: string; staff: string }>,
): { col: number; subject: Subject }[] {
  const subjectCols: { col: number; subject: Subject }[] = [];
  const seenCodes = new Set<string>();

  for (let col = 0; col < header.length; col++) {
    if (skipCols.has(col)) continue;
    const text = cellStr(header[col]);
    if (!text) continue;
    if (isIdentityHeader(text)) continue;

    const extracted = extractSubjectFromHeader(text);
    if (!extracted) continue;
    if (seenCodes.has(extracted.code)) continue;
    seenCodes.add(extracted.code);

    const fromFront = frontSubjects.get(extracted.code);
    subjectCols.push({
      col,
      subject: {
        code: extracted.code,
        name:
          extracted.name !== extracted.code
            ? extracted.name
            : fromFront?.name || extracted.name,
        staff: fromFront?.staff ?? "",
      },
    });
  }

  return subjectCols.sort((a, b) => a.col - b.col);
}

export function parseWorkbook(data: ArrayBuffer): ParsedSheet {
  const wb = XLSX.read(data, { type: "array", cellDates: true, dense: false });

  const markName =
    wb.SheetNames.find((n) => /^mark$/i.test(n)) ??
    wb.SheetNames.find((n) => /mark/i.test(n)) ??
    wb.SheetNames[wb.SheetNames.length - 1];

  const markSheet = markName ? wb.Sheets[markName] : undefined;
  if (!markSheet) throw new Error("No worksheet found in the file");

  // Force full used range so trailing subject columns are not dropped
  const ref = markSheet["!ref"] || "A1";
  const markGrid = XLSX.utils.sheet_to_json<unknown[]>(markSheet, {
    header: 1,
    defval: "",
    raw: true,
    blankrows: false,
    range: ref,
  });

  // Normalize row lengths to the widest row
  const maxWidth = markGrid.reduce(
    (m, row) => Math.max(m, (row as unknown[]).length),
    0,
  );
  for (const row of markGrid) {
    const r = row as unknown[];
    while (r.length < maxWidth) r.push("");
  }

  const meta = parseMetaFromGrid(markGrid);

  const frontName =
    wb.SheetNames.find((n) => /^front$/i.test(n)) ??
    wb.SheetNames.find((n) => /front/i.test(n)) ??
    wb.SheetNames[0];

  let frontSubjects = new Map<string, { name: string; staff: string }>();
  if (frontName && wb.Sheets[frontName] && frontName !== markName) {
    const frontGrid = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[frontName], {
      header: 1,
      defval: "",
      raw: true,
    });
    const frontMeta = parseMetaFromGrid(frontGrid);
    if (frontMeta.institution) meta.institution = frontMeta.institution;
    if (frontMeta.department) meta.department = frontMeta.department;
    if (frontMeta.title) meta.title = frontMeta.title;
    if (frontMeta.year) meta.year = frontMeta.year;
    if (frontMeta.semester) meta.semester = frontMeta.semester;
    if (frontMeta.batch) meta.batch = frontMeta.batch;
    if (frontMeta.section) meta.section = frontMeta.section;
    frontSubjects = parseSubjectsFromFront(frontGrid);
  }

  // Find every REG. NO. header row and pick the one with the most subjects
  // (main list has all subjects; arrear section headers may be shorter)
  const candidateIdxs: number[] = [];
  markGrid.forEach((row, i) => {
    if (
      row.some(
        (c) => typeof c === "string" && /reg\.?\s*no/i.test(String(c)),
      )
    ) {
      candidateIdxs.push(i);
    }
  });
  if (!candidateIdxs.length) {
    throw new Error("Could not find a REG. NO. header row");
  }

  let headerIdx = candidateIdxs[0];
  let subjectCols: { col: number; subject: Subject }[] = [];
  let colReg = 1;
  let colName = 2;
  let colGender = 3;
  let colQuota = 4;
  let colStay = 5;

  for (const idx of candidateIdxs) {
    const header = markGrid[idx] as unknown[];

    let cReg = 1;
    let cName = 2;
    let cGender = 3;
    let cQuota = 4;
    let cStay = 5;

    header.forEach((cell, col) => {
      const h = cellStr(cell).toUpperCase();
      if (/REG/.test(h) && /NO/.test(h)) cReg = col;
      if (/NAME/.test(h) && !/SUB/.test(h) && !CODE_RE.test(h)) cName = col;
      if (/^B\s*\/?\s*G$/.test(h) || h === "B/G") cGender = col;
      if (/^C\s*\/?\s*M$/.test(h) || h === "C/M" || h === "G/M") cQuota = col;
      if (/^H\s*\/?\s*DS$/.test(h) || h === "H/DS") cStay = col;
    });

    const skip = new Set([cReg, cName, cGender, cQuota, cStay]);
    // Also skip E/T if present
    header.forEach((cell, col) => {
      if (/^E\s*\/?\s*T$/.test(cellStr(cell).toUpperCase())) skip.add(col);
    });

    const found = collectSubjectsFromRow(header, skip, frontSubjects);

    // Fallback: columns after H/DS / E/T until TOTAL that have mark-like data
    if (found.length < 3) {
      let startCol = Math.max(cStay + 1, cGender + 1, cQuota + 1, cName + 1);
      // skip E/T
      if (/^E\s*\/?\s*T$/i.test(cellStr(header[startCol]))) startCol += 1;

      for (let col = startCol; col < header.length; col++) {
        const text = cellStr(header[col]);
        if (!text) continue;
        if (/^(total|pass|arrear)/i.test(text)) break;
        if (isIdentityHeader(text)) continue;
        const extracted = extractSubjectFromHeader(text);
        if (extracted && !found.some((f) => f.subject.code === extracted.code)) {
          const fromFront = frontSubjects.get(extracted.code);
          found.push({
            col,
            subject: {
              code: extracted.code,
              name:
                extracted.name !== extracted.code
                  ? extracted.name
                  : fromFront?.name || extracted.name,
              staff: fromFront?.staff ?? "",
            },
          });
        }
      }
      found.sort((a, b) => a.col - b.col);
    }

    if (found.length > subjectCols.length) {
      subjectCols = found;
      headerIdx = idx;
      colReg = cReg;
      colName = cName;
      colGender = cGender;
      colQuota = cQuota;
      colStay = cStay;
    }
  }

  // If MARK still has few subjects, append any FRONT subjects as metadata only
  // (cannot invent mark columns without positions)
  if (subjectCols.length === 0 && frontSubjects.size > 0) {
    throw new Error(
      "Found subjects on FRONT sheet but no mark columns on MARK. Check the mark header row.",
    );
  }

  if (!subjectCols.length) {
    throw new Error(
      "No subject columns found. Header must include codes like CS3452, IT3401.",
    );
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
