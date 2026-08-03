import * as XLSX from "xlsx";
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

export function exportAnalysisSheet(
  a: Analysis,
  subjects: Subject[],
  meta: ExportMeta,
) {
  const wb = XLSX.utils.book_new();
  const cat = a.categoryCounts;
  const hostC = (cat["C-H-B"] ?? 0) + (cat["C-H-G"] ?? 0);
  const dayC = (cat["C-DS-B"] ?? 0) + (cat["C-DS-G"] ?? 0);
  const hostM = (cat["M-H-B"] ?? 0) + (cat["M-H-G"] ?? 0);
  const dayM = (cat["M-DS-B"] ?? 0) + (cat["M-DS-G"] ?? 0);

  const catCols = (p: (typeof a.particulars)[0]) => [
    p.byCategory["C-H-B"],
    p.byCategory["C-H-G"],
    p.byCategory["C-DS-B"],
    p.byCategory["C-DS-G"],
    p.byCategory["M-H-B"],
    p.byCategory["M-H-G"],
    p.byCategory["M-DS-B"],
    p.byCategory["M-DS-G"],
  ];

  const front: (string | number)[][] = [
    [meta.institution],
    [meta.department],
    [meta.title],
    [
      "YEAR  :",
      meta.year,
      "SEMESTER:",
      meta.semester,
      "BATCH:",
      "",
      "BATCH:",
      meta.batch,
      "",
      "",
      "SECTION:",
      "",
      "",
      meta.section,
    ],
    [],
    [
      "S.No.",
      "PARTICULARS",
      "",
      "",
      "",
      "",
      "",
      "TOTAL",
      "COUNSELLING",
      "",
      "",
      "",
      "MANAGEMENT",
      "",
      "",
      "",
    ],
    [
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      `HOSTELLER(${hostC})`,
      "",
      `DAY SCHOLAR(${dayC})`,
      "",
      `HOSTELLER(${hostM})`,
      "",
      `DAY SCHOLAR(${dayM})`,
      "",
    ],
    [
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "BOYS",
      "GIRLS",
      "BOYS",
      "GIRLS",
      "BOYS",
      "GIRLS",
      "BOYS",
      "GIRLS",
    ],
    ...a.particulars.map((p, i) => [
      i + 1,
      p.label,
      "",
      "",
      "",
      "",
      "",
      p.total,
      ...catCols(p),
    ]),
    [
      a.particulars.length + 1,
      "PASS PERCENTAGE",
      "",
      "",
      "",
      "",
      "",
      Number(fmt(a.passPercent, 2)),
    ],
    [],
    ["TOPPER LIST"],
    [
      "",
      "S.NO",
      "REG.NO",
      "PHOTO",
      "STUDENTS NAME",
      "",
      "",
      "",
      "",
      "TOTAL",
      "",
      "PASS %",
      "",
      "RANK",
    ],
    ...a.toppers.slice(0, 3).map((r, i) => {
      const rankLabel =
        r.rank === 1
          ? "I"
          : r.rank === 2
            ? "II"
            : r.rank === 3
              ? "III"
              : String(r.rank);
      return [
        "",
        i + 1,
        r.student.reg,
        "",
        r.student.name,
        "",
        "",
        "",
        "",
        r.total,
        "",
        Number(fmt(r.percent, 2)),
        "",
        rankLabel,
      ];
    }),
    [],
    ["SUBJECT WISE PERFORMANCE"],
    [
      "",
      "S.NO",
      "SUB CODE",
      "SUBJECT NAME",
      "",
      "STAFF NAME",
      "",
      "",
      "",
      "APPEARED",
      "ABSENT",
      "PASSED",
      "FAILED",
      "PASS %",
    ],
    ...a.subjectStats.map((s, i) => [
      "",
      i + 1,
      s.subject.code,
      s.subject.name,
      "",
      s.subject.staff || "",
      "",
      "",
      "",
      s.appeared + s.absent,
      s.absent,
      s.passed,
      s.failed,
      Math.round(s.passPercent),
    ]),
    ...Array.from({ length: 20 }, () => [] as (string | number)[]),
    [
      "",
      "PREPARED BY",
      "",
      "",
      "HOD",
      "",
      "",
      "",
      "",
      "VICE PRINCIPAL",
      "",
      "",
      "",
      "PRINCIPAL",
    ],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(front), "FRONT");

  const subjectHeaders = subjects.map((s) => `${s.code} & ${s.name}`);
  const markHeader = [
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

  const rowFor = (r: (typeof a.results)[0], i: number) => [
    i + 1,
    r.student.reg,
    r.student.name,
    r.student.gender,
    r.student.quota,
    r.student.stay,
    "E",
    ...r.student.marks.map((m) => (m === null ? "" : m)),
    r.total,
    Number(fmt(r.percent, 2)),
    r.arrears,
  ];

  const sorted = [...a.results].sort((x, y) =>
    x.student.reg.localeCompare(y.student.reg),
  );

  const allClear = sorted.filter((r) => r.arrears === 0);
  const oneArrear = sorted.filter((r) => r.arrears === 1);
  const twoArrear = sorted.filter((r) => r.arrears === 2);
  const threePlus = sorted.filter((r) => r.arrears >= 3);

  const markSheet: (string | number)[][] = [
    [meta.institution],
    [meta.department],
    ["ASSESSMENT RESULT ANALYSIS"],
    [
      "YEAR  :",
      meta.year,
      "SEMESTER:",
      meta.semester,
      "BATCH:",
      "",
      "BATCH:",
      meta.batch,
      "",
      "",
      "",
      "SECTION:",
      "",
      "",
      meta.section,
    ],
    [],
    markHeader,
    ...sorted.map((r, i) => rowFor(r, i)),
    [
      "Total No. of Students",
      "",
      "",
      "",
      "",
      "",
      "",
      ...subjects.map(() => a.totalStudents),
    ],
    [
      "No. of Students Present",
      "",
      "",
      "",
      "",
      "",
      "",
      ...a.subjectStats.map((s) => s.appeared),
    ],
    [
      "No. of Students Absent",
      "",
      "",
      "",
      "",
      "",
      "",
      ...a.subjectStats.map((s) => s.absent),
    ],
    [
      "No. of Students Pass",
      "",
      "",
      "",
      "",
      "",
      "",
      ...a.subjectStats.map((s) => s.passed),
    ],
    [
      "No. of Students fail",
      "",
      "",
      "",
      "",
      "",
      "",
      ...a.subjectStats.map((s) => s.failed),
    ],
    [
      "PASS%",
      "",
      "",
      "",
      "",
      "",
      "",
      ...a.subjectStats.map((s) => Number(fmt(s.passPercent, 2))),
    ],
    [],
    [meta.institution],
    [meta.department],
    ["ASSESSMENT RESULT ANALYSIS"],
    ["ALL CLEAR STUDENTS"],
    markHeader,
    ...allClear.map((r, i) => rowFor(r, i)),
    ["ONE SUBJECTS ARREAR STUDENTS"],
    markHeader,
    ...oneArrear.map((r, i) => rowFor(r, i)),
    ["TWO SUBJECTS ARREAR STUDENTS"],
    markHeader,
    ...twoArrear.map((r, i) => rowFor(r, i)),
    ["THREE & ABOVE SUBJECTS ARREAR STUDENTS"],
    markHeader,
    ...threePlus.map((r, i) => rowFor(r, i)),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(markSheet), "MARK");

  const safeBatch = (meta.batch || "batch").replace(/[^\w-]+/g, "_");
  const safeSection = (meta.section || "sec").replace(/[^\w-]+/g, "_");
  XLSX.writeFile(wb, `Result_Analysis_${safeBatch}_${safeSection}.xlsx`);
}
