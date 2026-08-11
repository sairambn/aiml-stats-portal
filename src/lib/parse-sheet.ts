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
  if (v instanceof Date) {
    const d = v.getDate().toString().padStart(2, "0");
    const m = (v.getMonth() + 1).toString().padStart(2, "0");
    const y = v.getFullYear().toString().slice(-2);
    return `${d}/${m}/${y}`;
  }
  return String(v).trim();
}

function parseMark(v: unknown): Mark {
  if (v === undefined || v === null || v === "") return "AB";
  if (typeof v === "number") {
    if (Number.isNaN(v)) return "AB";
    return Math.round(v);
  }
  const s = String(v).trim();
  if (!s) return "AB";
  if (/^(ab|od|na|n\/?a|-)$/i.test(s)) return "AB";
  const n = Number(s);
  if (!Number.isFinite(n)) return "AB";
  return Math.round(n);
}

const CODE_RE = /([A-Z]{2,4}\d{3,4})/i;

const SHORT_NAMES: Record<string, string> = {
  TOC: "Theory of Computation",
  OS: "Operating Systems",
  DBDM: "Database Design and Management",
  DBMS: "Database Management Systems",
  AI: "Artificial Intelligence",
  ML: "Machine Learning",
  MI: "Machine Learning",
  EVS: "Environmental Sciences and Sustainability",
  WE: "Web Essentials",
  AIML: "Artificial Intelligence and Machine Learning",
  DM: "Discrete Mathematics",
  OOSE: "Object Oriented Software Engineering",
  DS: "Data Structures",
  JP: "Java Programming",
  NLP: "Natural Language Processing",
  DL: "Deep Learning",
  DC: "Data Communication",
  CCS: "Cloud Computing Systems",
  CC: "Cloud Computing",
  STA: "Statistics and Data Analytics",
  CNS: "Cryptography and Network Security",
  SE: "Software Engineering",
  CN: "Computer Networks",
  DAA: "Design and Analysis of Algorithms",
  PYTHON: "Python Programming",
  R: "R Programming",
};

function parseMetaFromGrid(grid: unknown[][]): Partial<ExportMeta> {
  const meta: Partial<ExportMeta> = {};
  if (!grid.length) return meta;
  for (let i = 0; i < Math.min(15, grid.length); i++) {
    const row = (grid[i] ?? []) as unknown[];
    const first = cellStr(row[0]);
    const joined = row.map(cellStr).filter(Boolean).join(" ");
    if (!meta.institution && first && /nagar|chennai|college|institute|university|road|salai/i.test(first)) {
      meta.institution = first;
    }
    if (!meta.department && /department/i.test(first || joined)) {
      meta.department = first || joined.match(/department[^,]*/i)?.[0] || joined;
    }
    if (!meta.title && first && /assessment|result\s*analysis|internal|mark\s*sheet|iat/i.test(first) && !/department/i.test(first)) {
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

function collectCodeInfo(grid: unknown[][]): Map<string, { name: string; staff: string }> {
  const map = new Map<string, { name: string; staff: string }>();
  for (const row of grid) {
    const cells = row as unknown[];
    for (let c = 0; c < cells.length; c++) {
      const raw = cellStr(cells[c]);
      const pure = raw.match(/^([A-Z]{2,4}\d{3,4})$/i);
      if (pure) {
        const code = pure[1].toUpperCase();
        let name = "";
        let staff = "";
        for (let k = c + 1; k < Math.min(c + 8, cells.length); k++) {
          const s = cellStr(cells[k]);
          if (!s) continue;
          if (!name && !/^\d+(\.\d+)?$/.test(s) && !CODE_RE.test(s) && !/\b(AP|PROF|DR\.?|MR\.?|MS\.?|MRS\.?)\b/i.test(s)) {
            name = s.replace(/^\s*&\s*/, "").trim();
            continue;
          }
          if (/\b(AP|PROF|DR\.?|MR\.?|MS\.?|MRS\.?)\b/i.test(s) || /\//.test(s)) {
            staff = s;
            break;
          }
        }
        const prev = map.get(code);
        map.set(code, { name: name || prev?.name || "", staff: staff || prev?.staff || "" });
        continue;
      }
      const combined = raw.match(/^([A-Z]{2,4}\d{3,4})\s*[&\-–:]\s*(.+)$/i);
      if (combined) {
        const code = combined[1].toUpperCase();
        const name = combined[2].trim();
        const prev = map.get(code);
        if (name && name.length > 3) map.set(code, { name, staff: prev?.staff || "" });
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
    /^NAME$/.test(u) ||
    (/NAME/.test(u) && !/SUB/.test(u) && !CODE_RE.test(u)) ||
    /^(B\/?G|C\/?M|G\/?M|H\/?DS|E\/?T)$/.test(u) ||
    /^(TOTAL|PASS\s*%?|ARREAR.*|ABSENT|STATUS|PERCENTAGE)$/i.test(u)
  );
}

function isDateLikeHeader(text: string): boolean {
  const t = text.trim().split(/[\s\n]+/)[0] || "";
  return (/^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}$/.test(t) || /^\d{4}-\d{2}-\d{2}/.test(t));
}

function extractShortCodeFromMixed(text: string): string | null {
  const parts = text.trim().split(/[\s\n\/]+/).filter(Boolean);
  for (const p of parts) {
    if (/^[A-Z]{2,6}$/i.test(p) && !/^\d/.test(p)) return p.toUpperCase();
  }
  return null;
}

function isShortSubjectCode(text: string): boolean {
  return /^[A-Z]{2,6}$/i.test(text.trim());
}

function extractSubjectFromHeader(text: string): { code: string; name: string } | null {
  if (!text || isIdentityHeader(text)) return null;
  const pure = text.match(/^\s*([A-Z]{2,4}\d{3,4})\s*$/i);
  if (pure) return { code: pure[1].toUpperCase(), name: pure[1].toUpperCase() };
  const m = text.match(/^\s*([A-Z]{2,4}\d{3,4})\s*(?:\(([^)]*)\)|[&\-–:]*\s*(.*))?\s*$/i);
  if (m) {
    const code = m[1].toUpperCase();
    let name = (m[2] || m[3] || "").trim().replace(/^[&\-–:]+\s*/, "");
    name = name.replace(/^\(|\)$/g, "").trim();
    if (/^(total|pass\s*%?|arrear.*|absent|status|percentage)$/i.test(name)) return null;
    return { code, name: name || code };
  }
  const anywhere = text.match(/\b([A-Z]{2,4}\d{3,4})\b/i);
  if (anywhere) {
    const code = anywhere[1].toUpperCase();
    const name = text.replace(anywhere[0], "").replace(/^[(&\-–:\s]+|[)\s]+$/g, "").trim();
    if (/^(total|pass|arrear|absent|status|percentage)$/i.test(name)) return { code, name: code };
    return { code, name: name || code };
  }
  return null;
}

function resolveSubjectName(code: string, extractedName: string, info: Map<string, { name: string; staff: string }>): string {
  const fromInfo = info.get(code)?.name;
  if (fromInfo && fromInfo.length > 3) return fromInfo;
  const short = (extractedName || code).toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (SHORT_NAMES[short]) return SHORT_NAMES[short];
  if (SHORT_NAMES[code.toUpperCase()]) return SHORT_NAMES[code.toUpperCase()];
  if (extractedName && extractedName !== code && extractedName.length > 3) return extractedName;
  return extractedName || code;
}

function findLegendStart(header: unknown[], afterCol: number): number | undefined {
  let emptyRun = 0;
  for (let col = afterCol; col < header.length; col++) {
    if (!cellStr(header[col])) {
      emptyRun += 1;
      if (emptyRun >= 2) return col - emptyRun + 1;
    } else emptyRun = 0;
  }
  return undefined;
}

function collectSubjectsFromRow(header: unknown[], skipCols: Set<number>, codeInfo: Map<string, { name: string; staff: string }>, preferBeforeCol?: number): { col: number; subject: Subject }[] {
  const subjectCols: { col: number; subject: Subject }[] = [];
  const seenCodes = new Set<string>();
  for (let col = 0; col < header.length; col++) {
    if (skipCols.has(col)) continue;
    if (preferBeforeCol != null && col >= preferBeforeCol) continue;
    const text = cellStr(header[col]);
    if (!text) continue;
    let extracted = extractSubjectFromHeader(text);
    if (!extracted && !isIdentityHeader(text)) {
      const short = extractShortCodeFromMixed(text);
      if (short) extracted = { code: short, name: short };
    }
    if (!extracted) {
      const pureCode = text.match(/^([A-Z]{2,4}\d{3,4})$/i);
      if (pureCode) extracted = { code: pureCode[1].toUpperCase(), name: pureCode[1].toUpperCase() };
    }
    if (!extracted || seenCodes.has(extracted.code)) continue;
    seenCodes.add(extracted.code);
    const info = codeInfo.get(extracted.code);
    subjectCols.push({ col, subject: { code: extracted.code, name: resolveSubjectName(extracted.code, extracted.name, codeInfo), staff: info?.staff ?? "" } });
  }
  return subjectCols.sort((a, b) => a.col - b.col);
}

function collectSubjectsFromDateAndCodeRows(dateRow: unknown[], codeRow: unknown[], skipCols: Set<number>, codeInfo: Map<string, { name: string; staff: string }>): { col: number; subject: Subject }[] {
  const subjectCols: { col: number; subject: Subject }[] = [];
  const seenCodes = new Set<string>();
  const maxCol = Math.max(dateRow.length, codeRow.length);
  for (let col = 0; col < maxCol; col++) {
    if (skipCols.has(col)) continue;
    const dateText = cellStr(dateRow[col]);
    const codeText = cellStr(codeRow[col]);
    let code = "";
    let nameHint = "";
    const codeFromText = (t: string) => {
      const m = t.match(/\b([A-Z]{2,4}\d{3,4})\b/i);
      return m ? m[1].toUpperCase() : null;
    };
    if (codeText && codeFromText(codeText)) { code = codeFromText(codeText)!; nameHint = code; }
    else if (dateText && codeFromText(dateText)) { code = codeFromText(dateText)!; nameHint = code; }
    else if (codeText && isShortSubjectCode(codeText)) { code = codeText.toUpperCase(); nameHint = code; }
    else if (dateText && isShortSubjectCode(dateText)) { code = dateText.toUpperCase(); nameHint = code; }
    else if (dateText && isDateLikeHeader(dateText) && codeText && /^[A-Z0-9]{2,8}$/i.test(codeText)) { code = codeText.toUpperCase(); nameHint = code; }
    if (!code) {
      const mixed = codeFromText(dateText) || codeFromText(codeText) || extractShortCodeFromMixed(dateText) || extractShortCodeFromMixed(codeText);
      if (mixed) { code = mixed; nameHint = mixed; }
    }
    if (!code || seenCodes.has(code)) continue;
    seenCodes.add(code);
    const info = codeInfo.get(code);
    subjectCols.push({ col, subject: { code, name: resolveSubjectName(code, nameHint, codeInfo), staff: info?.staff ?? "" } });
  }
  return subjectCols.sort((a, b) => a.col - b.col);
}

function isStudentReg(reg: string): boolean {
  if (!reg || reg.length < 5) return false;
  if (/^\d{6,}$/.test(reg)) return true;
  if (/^[0-9A-Z]{6,}$/i.test(reg) && /\d/.test(reg)) return true;
  return false;
}

export function parseWorkbook(data: ArrayBuffer): ParsedSheet {
  const wb = XLSX.read(data, { type: "array", cellDates: true, dense: false });
  const markName = wb.SheetNames.find((n) => /^mark$/i.test(n)) ?? wb.SheetNames.find((n) => /mark/i.test(n)) ?? wb.SheetNames[wb.SheetNames.length - 1];
  const markSheet = markName ? wb.Sheets[markName] : undefined;
  if (!markSheet) throw new Error("No worksheet found in the file");
  const ref = markSheet["!ref"] || "A1";
  const markGrid = XLSX.utils.sheet_to_json<unknown[]>(markSheet, { header: 1, defval: "", raw: true, blankrows: false, range: ref });
  const maxWidth = markGrid.reduce((m, row) => Math.max(m, (row as unknown[]).length), 0);
  for (const row of markGrid) { const r = row as unknown[]; while (r.length < maxWidth) r.push(""); }
  const meta = parseMetaFromGrid(markGrid);
  let codeInfo = collectCodeInfo(markGrid);
  const frontName = wb.SheetNames.find((n) => /^front$/i.test(n)) ?? wb.SheetNames.find((n) => /front/i.test(n));
  if (frontName && wb.Sheets[frontName] && frontName !== markName) {
    const frontGrid = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[frontName], { header: 1, defval: "", raw: true });
    const frontMeta = parseMetaFromGrid(frontGrid);
    if (frontMeta.institution) meta.institution = frontMeta.institution;
    if (frontMeta.department) meta.department = frontMeta.department;
    if (frontMeta.title) meta.title = frontMeta.title;
    if (frontMeta.year) meta.year = frontMeta.year;
    if (frontMeta.semester) meta.semester = frontMeta.semester;
    if (frontMeta.batch) meta.batch = frontMeta.batch;
    if (frontMeta.section) meta.section = frontMeta.section;
    const frontInfo = collectCodeInfo(frontGrid);
    for (const [code, info] of frontInfo) {
      const prev = codeInfo.get(code);
      codeInfo.set(code, { name: info.name || prev?.name || "", staff: info.staff || prev?.staff || "" });
    }
  }
  const candidateIdxs: number[] = [];
  markGrid.forEach((row, i) => {
    if (row.some((c) => {
      const raw = String(c ?? "");
      const s = raw.toLowerCase().replace(/[.\s_\-]/g, "");
      return (
        /reg.*n[ou]/.test(s) ||
        s.includes("regno") ||
        s.includes("registernumber") ||
        s.includes("registrationno") ||
        s.includes("registrationnumber")
      );
    })) candidateIdxs.push(i);
  });
  if (!candidateIdxs.length) throw new Error("Could not find a REG. NO. / Reg.No / Register No. header row");
  let headerIdx = candidateIdxs[0];
  let subjectCols: { col: number; subject: Subject }[] = [];
  let colReg = 1;
  let colName = 2;
  let colGender = -1;
  let colQuota = -1;
  let colStay = -1;
  let dataStartOffset = 1;
  for (const idx of candidateIdxs) {
    const header = markGrid[idx] as unknown[];
    let cReg = -1, cName = -1, cGender = -1, cQuota = -1, cStay = -1;
    header.forEach((cell, col) => {
      const h = cellStr(cell).toUpperCase().replace(/[\s.\-_]/g, "");
      if (/REG.*N[OU]/.test(h) || h === "REGNO" || h.includes("REGISTERNO") || h.includes("REGISTERNUMBER") || h.includes("REGISTRATIONNO") || h.includes("REGISTRATIONNUMBER")) cReg = col;
      if ((h === "NAME" || h === "STUDENT'SNAME" || h === "STUDENTSNAME" || h === "NAMEOFTHESTUDENT" || h === "STUDENTNAME" || h === "NAMEOFSTUDENT") && !CODE_RE.test(h)) cName = col;
      if (h === "B/G" || h === "BG" || h === "GENDER") cGender = col;
      if (h === "C/M" || h === "CM" || h === "G/M" || h === "QUOTA") cQuota = col;
      if (h === "H/DS" || h === "HDS" || h === "HOSTEL") cStay = col;
    });
    if (cReg < 0) cReg = 1;
    if (cName < 0) cName = 2;
    const skip = new Set<number>([cReg, cName]);
    if (cGender >= 0) skip.add(cGender);
    if (cQuota >= 0) skip.add(cQuota);
    if (cStay >= 0) skip.add(cStay);
    header.forEach((cell, col) => {
      const h = cellStr(cell).toUpperCase().replace(/\s+/g, "");
      if (/^(SL\.?NO\.?|S\.?NO\.?|SNO|S\.NO)$/.test(h)) skip.add(col);
      if (/^E\/?T$/.test(h)) skip.add(col);
    });
    const legendStart = findLegendStart(header, Math.max(cReg, cName, cGender, cQuota, cStay, 0) + 1);
    let found = collectSubjectsFromRow(header, skip, codeInfo, legendStart);
    let offset = 1;
    if (found.length === 0 && idx + 1 < markGrid.length) {
      const nextRow = markGrid[idx + 1] as unknown[];
      const fromDateCode = collectSubjectsFromDateAndCodeRows(header, nextRow, skip, codeInfo);
      if (fromDateCode.length > 0) { found = fromDateCode; offset = 2; }
    }
    if (found.length === 0) {
      const shortOnly: { col: number; subject: Subject }[] = [];
      const seen = new Set<string>();
      for (let col = 0; col < header.length; col++) {
        if (skip.has(col)) continue;
        const t = cellStr(header[col]);
        if (!isShortSubjectCode(t)) continue;
        const code = t.toUpperCase();
        if (seen.has(code)) continue;
        seen.add(code);
        shortOnly.push({ col, subject: { code, name: resolveSubjectName(code, code, codeInfo), staff: codeInfo.get(code)?.staff ?? "" } });
      }
      if (shortOnly.length > 0) found = shortOnly;
    }
    if (found.length > subjectCols.length) {
      subjectCols = found;
      headerIdx = idx;
      colReg = cReg; colName = cName; colGender = cGender; colQuota = cQuota; colStay = cStay;
      dataStartOffset = offset;
    }
  }

  // Ultimate fallback: scan the header row and the next 2 rows for any subject-code-like values
  if (!subjectCols.length) {
    const scanRows = [headerIdx, headerIdx + 1, headerIdx + 2].filter(i => i < markGrid.length);
    const skip = new Set<number>([colReg, colName]);
    const seen = new Set<string>();
    const found: { col: number; subject: Subject }[] = [];
    for (const rIdx of scanRows) {
      const row = markGrid[rIdx] as unknown[];
      for (let col = 0; col < row.length; col++) {
        if (skip.has(col)) continue;
        const text = cellStr(row[col]);
        if (!text) continue;
        const m = text.match(/\b([A-Z]{2,4}\d{3,4})\b/i);
        if (m) {
          const code = m[1].toUpperCase();
          if (seen.has(code)) continue;
          seen.add(code);
          found.push({ col, subject: { code, name: resolveSubjectName(code, code, codeInfo), staff: codeInfo.get(code)?.staff ?? "" } });
          continue;
        }
        if (/^[A-Z]{2,6}$/i.test(text.trim())) {
          const code = text.trim().toUpperCase();
          if (seen.has(code)) continue;
          seen.add(code);
          found.push({ col, subject: { code, name: resolveSubjectName(code, code, codeInfo), staff: codeInfo.get(code)?.staff ?? "" } });
        }
      }
    }
    if (found.length > 0) {
      subjectCols = found.sort((a, b) => a.col - b.col);
      dataStartOffset = 2;
    }
  }
  if (!subjectCols.length) throw new Error("No subject columns found. Header must include codes like CS3452, AL3501, CCS335, GE3451 or short codes (NLP, DM, OS).");
  const parsed: Student[] = [];
  const seen = new Set<string>();
  for (const row of markGrid.slice(headerIdx + dataStartOffset)) {
    const cells = row as unknown[];
    const first = cellStr(cells[0]).toUpperCase();
    const second = cellStr(cells[1]).toUpperCase();
    if (/total\s*no/i.test(first) || /total\s*no/i.test(second) || /no\.?\s*of\s*students/i.test(first) || /no\.?\s*of\s*students/i.test(second) || /no\.?\s*of\s*pass/i.test(first) || /no\.?\s*of\s*pass/i.test(second) || /no\.?\s*of\s*fail/i.test(first) || /no\.?\s*of\s*fail/i.test(second) || /percentage/i.test(first) || /percentage/i.test(second) || /all\s*clear/i.test(first) || /one\s*subject/i.test(first) || /two\s*subject/i.test(first) || /three/i.test(first) || /^arrear/i.test(first) || /^pass%/i.test(first) || /jeppiaar/i.test(first) || /department/i.test(first) || /assessment/i.test(first) || /hostel/i.test(first) || /hostel/i.test(second)) break;
    const reg = cellStr(cells[colReg]).replace(/\.0$/, "");
    if (!isStudentReg(reg) || seen.has(reg)) continue;
    seen.add(reg);
    const genderRaw = colGender >= 0 ? cellStr(cells[colGender]).toUpperCase() : "";
    const quotaRaw = colQuota >= 0 ? cellStr(cells[colQuota]).toUpperCase() : "";
    const stayRaw = colStay >= 0 ? cellStr(cells[colStay]).toUpperCase() : "";
    parsed.push({ reg, name: cellStr(cells[colName]), gender: genderRaw.startsWith("G") ? "G" : "B", quota: quotaRaw.startsWith("M") ? "M" : "C", stay: stayRaw.startsWith("H") ? "H" : "DS", marks: subjectCols.map(({ col }) => parseMark(cells[col])) });
  }
  if (!parsed.length) throw new Error("No student rows found in the mark sheet");
  return { students: parsed, subjects: subjectCols.map((s) => s.subject), meta };
}
