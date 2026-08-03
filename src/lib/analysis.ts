export type Mark = number | "AB" | null;

export type Student = {
  reg: string;
  name: string;
  gender: string; // B | G
  quota: string; // C (counselling) | M (management)
  stay: string; // H (hosteller) | DS (day scholar)
  marks: Mark[];
};

export type Subject = { code: string; name: string; staff: string };

export const isAbsent = (m: Mark) => typeof m === "string" && m.toUpperCase() === "AB";
export const numeric = (m: Mark): number | null => (typeof m === "number" ? m : null);

export type StudentResult = {
  student: Student;
  total: number;
  percent: number;
  arrears: number;
  absents: number;
  status: "PASS" | "FAIL";
  rank: number;
};

export type SubjectStat = {
  subject: Subject;
  index: number;
  appeared: number;
  absent: number;
  passed: number;
  failed: number;
  passPercent: number;
  average: number;
  highest: number;
};

export type CategoryKey =
  | "C-H-B"
  | "C-H-G"
  | "C-DS-B"
  | "C-DS-G"
  | "M-H-B"
  | "M-H-G"
  | "M-DS-B"
  | "M-DS-G";

export type ParticularRow = {
  label: string;
  total: number;
  byCategory: Record<CategoryKey, number>;
};

export const categoryKey = (s: Student): CategoryKey =>
  `${s.quota.toUpperCase().startsWith("M") ? "M" : "C"}-${
    s.stay.toUpperCase().startsWith("H") ? "H" : "DS"
  }-${s.gender.toUpperCase().startsWith("G") ? "G" : "B"}` as CategoryKey;

export const CATEGORY_KEYS: CategoryKey[] = [
  "C-H-B",
  "C-H-G",
  "C-DS-B",
  "C-DS-G",
  "M-H-B",
  "M-H-G",
  "M-DS-B",
  "M-DS-G",
];

const emptyCounts = () =>
  CATEGORY_KEYS.reduce(
    (acc, k) => ({ ...acc, [k]: 0 }),
    {} as Record<CategoryKey, number>,
  );

export function computeStudentResults(
  students: Student[],
  passMark: number,
): StudentResult[] {
  const rows = students.map((student) => {
    const nums = student.marks.map(numeric);
    const total = nums.reduce((a, b) => a + (b ?? 0), 0);
    const absents = student.marks.filter(isAbsent).length;
    const arrears =
      student.marks.filter((m) => isAbsent(m) || (numeric(m) ?? 0) < passMark).length;
    const count = student.marks.length || 1;
    return {
      student,
      total,
      percent: total / count,
      arrears,
      absents,
      status: (arrears === 0 ? "PASS" : "FAIL") as "PASS" | "FAIL",
      rank: 0,
    };
  });

  const sorted = [...rows].sort((a, b) => b.total - a.total);
  let lastTotal: number | null = null;
  let lastRank = 0;
  sorted.forEach((r, i) => {
    if (r.total === lastTotal) r.rank = lastRank;
    else {
      r.rank = i + 1;
      lastRank = r.rank;
      lastTotal = r.total;
    }
  });
  return rows;
}

export function computeSubjectStats(
  students: Student[],
  subjects: Subject[],
  passMark: number,
): SubjectStat[] {
  return subjects.map((subject, index) => {
    const marks = students.map((s) => s.marks[index] ?? null);
    const absent = marks.filter(isAbsent).length;
    const present = marks.map(numeric).filter((n): n is number => n !== null);
    const passed = present.filter((n) => n >= passMark).length;
    const failed = present.length - passed + absent;
    const appeared = present.length;
    return {
      subject,
      index,
      appeared,
      absent,
      passed,
      failed,
      passPercent: appeared ? (passed / appeared) * 100 : 0,
      average: appeared ? present.reduce((a, b) => a + b, 0) / appeared : 0,
      highest: appeared ? Math.max(...present) : 0,
    };
  });
}

export function computeParticulars(results: StudentResult[]): ParticularRow[] {
  const build = (label: string, predicate: (r: StudentResult) => boolean): ParticularRow => {
    const byCategory = emptyCounts();
    let total = 0;
    for (const r of results) {
      if (!predicate(r)) continue;
      total += 1;
      byCategory[categoryKey(r.student)] += 1;
    }
    return { label, total, byCategory };
  };

  return [
    build("Total no of students", () => true),
    build("No of students appeared", (r) => r.absents < r.student.marks.length),
    build("No of students passed in all subjects", (r) => r.arrears === 0),
    build("No of students failed", (r) => r.arrears > 0),
    build("No of students failed in one subject", (r) => r.arrears === 1),
    build("No of students failed in two subjects", (r) => r.arrears === 2),
    build("No of students failed in 3 & above subjects", (r) => r.arrears >= 3),
  ];
}

export type Analysis = {
  results: StudentResult[];
  subjectStats: SubjectStat[];
  particulars: ParticularRow[];
  totalStudents: number;
  appeared: number;
  passedAll: number;
  failed: number;
  totalAbsentEntries: number;
  studentsWithAbsence: number;
  passPercent: number;
  classAverage: number;
  toppers: StudentResult[];
  categoryCounts: Record<CategoryKey, number>;
};

export function analyse(
  students: Student[],
  subjects: Subject[],
  passMark: number,
): Analysis {
  const results = computeStudentResults(students, passMark);
  const subjectStats = computeSubjectStats(students, subjects, passMark);
  const particulars = computeParticulars(results);
  const appeared = results.filter((r) => r.absents < r.student.marks.length).length;
  const passedAll = results.filter((r) => r.arrears === 0).length;
  const categoryCounts = emptyCounts();
  for (const r of results) categoryCounts[categoryKey(r.student)] += 1;

  return {
    results,
    subjectStats,
    particulars,
    totalStudents: results.length,
    appeared,
    passedAll,
    failed: results.length - passedAll,
    totalAbsentEntries: results.reduce((a, r) => a + r.absents, 0),
    studentsWithAbsence: results.filter((r) => r.absents > 0).length,
    passPercent: results.length ? (passedAll / results.length) * 100 : 0,
    classAverage: results.length
      ? results.reduce((a, r) => a + r.percent, 0) / results.length
      : 0,
    toppers: [...results].sort((a, b) => a.rank - b.rank || a.student.name.localeCompare(b.student.name)).slice(0, 5),
    categoryCounts,
  };
}

export const fmt = (n: number, digits = 1) =>
  Number.isFinite(n) ? n.toFixed(digits) : "0";
