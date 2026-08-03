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
  // AB = absent, OD = on duty (treat as absent for scoring)
  if (/^(ab|od|na|n\/?a|-)$/i.test(s)) return "AB";
  const n = Number(s);
  return Number.isFinite(n) ? Math.round(n) : 0;
}

const CODE_RE = /([A-Z]{2,4}\d{3,4})/i;

function parseMetaFromGrid(grid: unknown[][]): Partial<ExportMeta> {
  const meta: Partial<ExportMeta> = {};
  if (!grid.length) return meta;

  for (let i = 0; i < Math.min(15, grid.length); i++) {
    const row = (grid[i] ?? []) as unknown[];
    const first = cellStr(row[0]);
    const joined = row.map(cellStr).filter(Boolean).join(" ");

    if (
      !meta.institution &&
      first &&
      /nagar|chennai|college|institute|university|road|salai/i.test(first)
    ) {
      meta.institution = first;
    }
    if (
      !meta.department &&
      (first || joined) &&
      /department/i.test(first || joined)
    ) {
      meta.department = first || joined;
    }
    if (
      !meta.title &&
      first &&
      /assessment|result\s*analysis|internal|mark\s*sheet|iat/i.test(first) &&
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
    row.some(
      (c) =>
        typeof c === "string" &&
        (/subject\s*wise/i.test(String(c)) ||
          (/subject\s*code/i.test(String(c)) &&
            row.some((x) => /subject\s*name/i.test(String(x ?? ""))))),
    ),
  );
  if (subjIdx < 0) return map;

  for (const row of grid.slice(subjIdx, subjIdx + 40)) {
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
    /^(SL\.?\s*NO\.?|S\.?NO\.?|SNO|S\.NO)$/.test(u) ||
    (/REG/.test(u) && /NO/.test(u)) ||
    (/^NAME$/.test(u) ||
      (/NAME/.test(u) && !/SUB/.test(u) && !CODE_RE.test(u))) ||
    /^(B\/?G|C\/?M|G\/?M|H\/?DS|E\/?T)$/.test(u) ||
    /^(TOTAL|PASS\s*%?|ARREAR.*|ABSENT|STATUS|PERCENTAGE)$/i.test(u)
  );
}

function extractSubjectFromHeader(
  text: string,
): { code: string; name: string } | null {
  if (!text || isIdentityHeader(text)) return null;

  // CS3452(TOC) or CS3452 & Theory... or CS3452-Name or CS3452 Name
  const m = text.match(
    /^\s*([A-Z]{2,4}\d{3,4})\s*(?:\(([^)]*)\)|[&\-–:]*\s*(.*))?\s*$/i,
  );
  if (m) {
    const code = m[1].toUpperCase();
    let name = (m[2] || m[3] || "").trim().replace(/^[&\-–:]+\s*/, "");
    // Strip surrounding parens leftovers
    name = name.replace(/^\(|\)$/g, "").trim();
    if (/^(total|pass\s*%?|arrear.*|absent|status|percentage)$/i.test(name))
      return null;
    return { code, name: name || code };
  }

  const anywhere = text.match(/\b([A-Z]{2,4}\d{3,4})\b/i);
  if (!anywhere) return null;
  const code = anywhere[1].toUpperCase();
  const name = text
    .replace(anywhere[0], "")
    .replace(/^[(&\-–:\s]+|[)\s]+$/g, "")
    .trim();
  if (/^(total|pass|arrear|absent|status|percentage)$/i.test(name)) return null;
  return { code, name: name || code };
}

function collectSubjectsFromRow(
  header: unknown[],
  skipCols: Set<number>,
  frontSubjects: Map<string, { name: string; staff: string }>,
  /** Prefer columns that look like mark columns (before long empty gap / legend) */
  preferBeforeCol?: number,
): { col: number; subject: Subject }[] {
  const subjectCols: { col: number; subject: Subject }[] = [];
  const seenCodes = new Set<string>();

  for (let col = 0; col < header.length; col++) {
    if (skipCols.has(col)) continue;
    if (preferBeforeCol != null && col >= preferBeforeCol) continue;
    const text = cellStr(header[col]);
    if (!text) continue;
    if (isIdentityHeader(text)) continue;

    const extracted = extractSubjectFromHeader(text);
    if (!extracted) continue;
    if (seenCodes.has(extracted.code)) continue;
    seenCodes.add(extracted.code);

    const fromFront = frontSubjects.get(extracted.code);
    let name = extracted.name;
    // Expand short names from legend/FRONT when header is only CODE(SHORT)
    if (
      fromFront?.name &&
      (name === extracted.code || name.length <= 6 || /^\(.*\)$/.test(name))
    ) {
      name = fromFront.name;
    }

    subjectCols.push({
      col,
      subject: {
        code: extracted.code,
        name,
        staff: fromFront?.staff ?? "",
      },
    });
  }

  return subjectCols.sort((a, b) => a.col - b.col);
}

/** Find first empty stretch after subjects that indicates a side legend */
function findLegendStart(header: unknown[], afterCol: number): number | undefined {
  let emptyRun = 0;
  for (let col = afterCol; col < header.length; col++) {
    if (!cellStr(header[col])) {
      emptyRun += 1;
      if (emptyRun >= 2) return col - emptyRun + 1;
    } else {
      emptyRun = 0;
    }
  }
  return undefined;
}

export function parseWorkbook(data: ArrayBuffer): ParsedSheet {
  const wb = XLSX.read(data, { type: "array", cellDates: true, dense: false });

  const markName =
    wb.SheetNames.find((n) => /^mark$/i.test(n)) ??
    wb.SheetNames.find((n) => /mark/i.test(n)) ??
    wb.SheetNames[wb.SheetNames.length - 1];

  const markSheet = markName ? wb.Sheets[markName] : undefined;
  if (!markSheet) throw new Error("No worksheet found in the file");

  const ref = markSheet["!ref"] || "A1";
  const markGrid = XLSX.utils.sheet_to_json<unknown[]>(markSheet, {
    header: 1,
    defval: "",
    raw: true,
    blankrows: false,
    range: ref,
  });

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
    wb.SheetNames.find((n) => /front/i.test(n));

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

  // Also parse subject legend from the same MARK sheet (code/name side columns)
  const markLegend = parseSubjectsFromFront(markGrid);
  for (const [code, info] of markLegend) {
    if (!frontSubjects.has(code)) frontSubjects.set(code, info);
  }

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
    throw new Error("Could not find a REG. NO. / Reg.No header row");
  }

  let headerIdx = candidateIdxs[0];
  let subjectCols: { col: number; subject: Subject }[] = [];
  // -1 means "column not present in this sheet"
  let colReg = 1;
  let colName = 2;
  let colGender = -1;
  let colQuota = -1;
  let colStay = -1;

  for (const idx of candidateIdxs) {
    const header = markGrid[idx] as unknown[];

    let cReg = -1;
    let cName = -1;
    let cGender = -1;
    let cQuota = -1;
    let cStay = -1;

    header.forEach((cell, col) => {
      const h = cellStr(cell).toUpperCase().replace(/\s+/g, "");
      if (/REG/.test(h) && /NO/.test(h)) cReg = col;
      if (
        (h === "NAME" || h === "STUDENT'SNAME" || h === "STUDENTSNAME") &&
        !CODE_RE.test(h)
      )
        cName = col;
      if (h === "B/G" || h === "BG") cGender = col;
      if (h === "C/M" || h === "CM" || h === "G/M") cQuota = col;
      if (h === "H/DS" || h === "HDS") cStay = col;
    });

    // Fallbacks only when clearly missing identity labels
    if (cReg < 0) cReg = 1;
    if (cName < 0) cName = 2;

    const skip = new Set<number>();
    skip.add(cReg);
    skip.add(cName);
    if (cGender >= 0) skip.add(cGender);
    if (cQuota >= 0) skip.add(cQuota);
    if (cStay >= 0) skip.add(cStay);
    // Skip S.No column if present
    header.forEach((cell, col) => {
      const h = cellStr(cell).toUpperCase();
      if (/^(SL\.?\s*NO\.?|S\.?NO\.?|SNO|S\.NO)$/.test(h.replace(/\s+/g, "")))
        skip.add(col);
      if (/^E\s*\/?\s*T$/.test(h)) skip.add(col);
    });

    // Don't pick up side legend codes (usually after empty columns)
    const legendStart = findLegendStart(
      header,
      Math.max(cReg, cName, cGender, cQuota, cStay, 0) + 1,
    );

    const found = collectSubjectsFromRow(
      header,
      skip,
      frontSubjects,
      legendStart,
    );

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

  if (!subjectCols.length) {
    throw new Error(
      "No subject columns found. Header must include codes like CS3452, AL3452, GE3451.",
    );
  }

  // Enrich short names from legend (e.g. TOC → Theory of Computation)
  for (const sc of subjectCols) {
    const info = frontSubjects.get(sc.subject.code);
    if (
      info?.name &&
      (sc.subject.name === sc.subject.code || sc.subject.name.length <= 6)
    ) {
      sc.subject.name = info.name;
    }
    if (info?.staff && !sc.subject.staff) sc.subject.staff = info.staff;
  }

  const parsed: Student[] = [];
  const seen = new Set<string>();

  for (const row of markGrid.slice(headerIdx + 1)) {
    const cells = row as unknown[];
    const first = cellStr(cells[0]).toUpperCase();

    if (
      /total\s*no/i.test(first) ||
      /no\.?\s*of\s*students/i.test(first) ||
      /no\.?\s*of\s*pass/i.test(first) ||
      /no\.?\s*of\s*fail/i.test(first) ||
      /percentage/i.test(first) ||
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

    const genderRaw =
      colGender >= 0 ? cellStr(cells[colGender]).toUpperCase() : "";
    const quotaRaw =
      colQuota >= 0 ? cellStr(cells[colQuota]).toUpperCase() : "";
    const stayRaw =
      colStay >= 0 ? cellStr(cells[colStay]).toUpperCase() : "";

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
