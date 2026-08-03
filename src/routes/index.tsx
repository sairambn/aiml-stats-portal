import { useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import * as XLSX from "xlsx";
import {
  analyse,
  fmt,
  isAbsent,
  type Mark,
  type Student,
  type Subject,
  CATEGORY_KEYS,
} from "@/lib/analysis";
import { meta as seedMeta, seedStudents, seedSubjects } from "@/data/seed";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "AIML Result Analysis Portal — Pass, Fail & Absentee Insights" },
      {
        name: "description",
        content:
          "Interactive internal assessment result analysis for the AIML department: pass counts, absentees, subject-wise performance, toppers and category-wise breakdown.",
      },
      { property: "og:title", content: "AIML Result Analysis Portal" },
      {
        property: "og:description",
        content:
          "Generate the full internal assessment result analysis sheet — pass, fail, absent and subject-wise stats — in seconds.",
      },
    ],
  }),
  component: Portal;
});

type TabKey = "overview" | "particulars" | "subjects" | "students" | "toppers";

const TABS: { key: TabKey; label: string }[] = [
  { key: "overview", label: "Overview" },
  { key: "particulars", label: "Particulars" },
  { key: "subjects", label: "Subject-wise" },
  { key: "students", label: "Students" },
  { key: "toppers", label: "Toppers" },
];

const CATEGORY_LABEL: Record<string, string> = {
  "C-H-B": "Counselling · Hosteller · Boys",
  "C-H-G": "Counselling · Hosteller · Girls",
  "C-DS-B": "Counselling · Day Scholar · Boys",
  "C-DS-G": "Counselling · Day Scholar · Girls",
  "M-H-B": "Management · Hosteller · Boys",
  "M-H-G": "Management · Hosteller · Girls",
  "M-DS-B": "Management · Day Scholar · Boys",
  "M-DS-G": "Management · Day Scholar · Girls",
};

function Portal() {
  const [students, setStudents] = useState<Student[]>(seedStudents);
  const [subjects, setSubjects] = useState<Subject[]>(seedSubjects);
  const [passMark, setPassMark] = useState(seedMeta.passMark);
  const [tab, setTab] = useState<TabKey>("overview");
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<"all" | "pass" | "fail" | "absent">("all");
  const [notice, setNotice] = useState<string | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  const a = useMemo(() => analyse(students, subjects, passMark), [students, subjects, passMark]);

  const visibleResults = useMemo(() => {
    const q = query.trim().toLowerCase();
    return a.results
      .filter((r) => {
        if (filter === "pass" && r.status !== "PASS") return false;
        if (filter === "fail" && r.status !== "FAIL") return false;
        if (filter === "absent" && r.absents === 0) return false;
        if (!q) return true;
        return (
          r.student.name.toLowerCase().includes(q) || r.student.reg.includes(q)
        );
      })
      .sort((x, y) => x.rank - y.rank);
  }, [a.results, query, filter]);

  function handleUpload(file: File) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target?.result, { type: "array" });
        const sheet =
          wb.Sheets["MARK"] ?? wb.Sheets[wb.SheetNames[wb.SheetNames.length - 1]];
        const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1 });
        const headerIdx = grid.findIndex((row) =>
          row.some((c) => typeof c === "string" && /reg\.?\s*no/i.test(c)),
        );
        if (headerIdx < 0) throw new Error("Could not find a 'REG. NO.' header row");
        const header = grid[headerIdx] as unknown[];
        const subjectCols: { col: number; subject: Subject }[] = [];
        header.forEach((cell, col) => {
          if (typeof cell !== "string" || col < 3) return;
          const m = cell.match(/^\s*([A-Z]{2}\d{4})\s*&?\s*(.*)$/);
          if (m) subjectCols.push({ col, subject: { code: m[1], name: m[2].trim(), staff: "" } });
        });
        if (!subjectCols.length) throw new Error("No subject columns found");

        const parsed: Student[] = [];
        for (const row of grid.slice(headerIdx + 1)) {
          const reg = String(row[1] ?? "").trim();
          if (!/^\d{6,}$/.test(reg)) continue;
          if (parsed.some((p) => p.reg === reg)) continue;
          parsed.push({
            reg,
            name: String(row[2] ?? "").trim(),
            gender: String(row[3] ?? "B").trim(),
            quota: String(row[4] ?? "C").trim(),
            stay: String(row[5] ?? "DS").trim(),
            marks: subjectCols.map(({ col }) => {
              const v = row[col];
              if (v === undefined || v === null || v === "") return null;
              if (typeof v === "number") return Math.round(v) as Mark;
              return /^ab$/i.test(String(v).trim()) ? ("AB" as Mark) : (Number(v) || 0);
            }),
          });
        }
        if (!parsed.length) throw new Error("No student rows found");
        setSubjects(subjectCols.map((s) => s.subject));
        setStudents(parsed);
        setNotice(`Loaded ${parsed.length} students and ${subjectCols.length} subjects from ${file.name}`);
      } catch (err) {
        setNotice(err instanceof Error ? err.message : "Could not read that file");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function exportSheet() {
    const wb = XLSX.utils.book_new();
    const front: (string | number)[][] = [
      [seedMeta.department],
      [seedMeta.title],
      [`YEAR: ${seedMeta.year}`, `SEMESTER: ${seedMeta.semester}`, `BATCH: ${seedMeta.batch}`, `SECTION: ${seedMeta.section}`],
      [],
      ["S.No.", "PARTICULARS", "TOTAL", ...CATEGORY_KEYS.map((k) => CATEGORY_LABEL[k])],
      ...a.particulars.map((p, i) => [
        i + 1,
        p.label.toUpperCase(),
        p.total,
        ...CATEGORY_KEYS.map((k) => p.byCategory[k]),
      ]),
      [a.particulars.length + 1, "PASS PERCENTAGE", Number(fmt(a.passPercent, 2))],
      [],
      ["SUBJECT WISE PERFORMANCE"],
      ["S.NO", "SUB CODE", "SUBJECT NAME", "STAFF NAME", "APPEARED", "ABSENT", "PASSED", "FAILED", "PASS %"],
      ...a.subjectStats.map((s, i) => [
        i + 1,
        s.subject.code,
        s.subject.name,
        s.subject.staff,
        s.appeared,
        s.absent,
        s.passed,
        s.failed,
        Number(fmt(s.passPercent, 2)),
      ]),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(front), "FRONT");

    const markSheet: (string | number)[][] = [
      ["SL.NO", "REG. NO.", "STUDENT'S NAME", "B/G", "C/M", "H/DS", ...subjects.map((s) => `${s.code} & ${s.name}`), "TOTAL", "PASS %", "ARREAR COUNT", "ABSENT", "STATUS"],
      ...[...a.results]
        .sort((x, y) => x.student.reg.localeCompare(y.student.reg))
        .map((r, i) => [
          i + 1,
          r.student.reg,
          r.student.name,
          r.student.gender,
          r.student.quota,
          r.student.stay,
          ...r.student.marks.map((m) => (m === null ? "" : m)),
          r.total,
          Number(fmt(r.percent, 2)),
          r.arrears,
          r.absents,
          r.status,
        ]),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(markSheet), "MARK");
    XLSX.writeFile(wb, `AIML_Result_Analysis_${seedMeta.section}.xlsx`);
  }

  return (
    <div className="min-h-screen bg-background text-foreground">
      <header className="border-b border-border bg-card">
        <div className="mx-auto max-w-7xl px-6 py-8">
          <p className="font-mono text-xs uppercase tracking-[0.28em] text-accent">
            {seedMeta.institution}
          </p>
          <h1 className="mt-3 font-display text-3xl leading-tight tracking-tight sm:text-4xl">
            AIML Result Analysis Portal
          </h1>
          <p className="mt-2 max-w-2xl text-sm text-muted-foreground">
            {seedMeta.title} · Year {seedMeta.year} · Semester {seedMeta.semester} ·
            Batch {seedMeta.batch} · Section {seedMeta.section}
          </p>

          <div className="mt-6 flex flex-wrap items-center gap-3">
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) handleUpload(f);
                e.target.value = "";
              }}
            />
            <button className="btn-primary" onClick={() => fileRef.current?.click()}>
              Upload mark sheet
            </button>
            <button className="btn-outline" onClick={exportSheet}>
              Export analysis (.xlsx)
            </button>
            <button className="btn-ghost" onClick={() => window.print()}>
              Print report
            </button>
            <label className="ml-auto flex items-center gap-2 text-sm text-muted-foreground">
              Pass mark
              <input
                type="number"
                min={1}
                max={100}
                value={passMark}
                onChange={(e) => setPassMark(Math.max(1, Math.min(100, Number(e.target.value) || 1)))}
                className="w-16 rounded-md border border-input bg-background px-2 py-1 text-right font-mono text-sm text-foreground"
              />
            </label>
          </div>
          {notice && (
            <p className="mt-4 rounded-md border border-accent/40 bg-accent/10 px-3 py-2 text-sm text-accent-foreground">
              {notice}
            </p>
          )}
        </div>
      </header>

      <nav className="sticky top-0 z-10 border-b border-border bg-card/95 backdrop-blur print:hidden">
        <div className="mx-auto flex max-w-7xl gap-1 overflow-x-auto px-6">
          {TABS.map((t) => (
            <button
              key={t.key}
              onClick={() => setTab(t.key)}
              className={`whitespace-nowrap border-b-2 px-4 py-3 text-sm font-medium transition-colors ${
                tab === t.key
                  ? "border-accent text-foreground"
                  : "border-transparent text-muted-foreground hover:text-foreground"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </nav>

      <main className="mx-auto max-w-7xl space-y-8 px-6 py-10">
        <section className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <Stat label="Total students" value={String(a.totalStudents)} hint={`${a.appeared} appeared`} />
          <Stat label="Passed all subjects" value={String(a.passedAll)} hint={`${fmt(a.passPercent)}% pass`} tone="good" />
          <Stat label="Failed" value={String(a.failed)} hint={`${a.results.filter((r) => r.arrears >= 3).length} with 3+ arrears`} tone="bad" />
          <Stat label="Absentees" value={String(a.totalAbsentEntries)} hint={`${a.studentsWithAbsence} students affected`} tone="warn" />
        </section>

        {tab === "overview" && (
          <div className="grid gap-6 lg:grid-cols-2">
            <Panel title="Pass percentage per subject">
              <div className="space-y-4">
                {a.subjectStats.map((s) => (
                  <div key={s.subject.code}>
                    <div className="flex items-baseline justify-between gap-3 text-sm">
                      <span className="font-medium">
                        <span className="font-mono text-xs text-accent">{s.subject.code}</span>{" "}
                        {s.subject.name}
                      </span>
                      <span className="font-mono text-xs text-muted-foreground">
                        {fmt(s.passPercent)}%
                      </span>
                    </div>
                    <div className="mt-2 h-2 overflow-hidden rounded-full bg-muted">
                      <div
                        className="h-full rounded-full bg-accent transition-[width] duration-500"
                        style={{ width: `${Math.min(100, s.passPercent)}%` }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            </Panel>
            <Panel title="Arrear distribution">
              <div className="space-y-3">
                {[
                  ["No arrears (all clear)", a.results.filter((r) => r.arrears === 0).length],
                  ["Failed in 1 subject", a.results.filter((r) => r.arrears === 1).length],
                  ["Failed in 2 subjects", a.results.filter((r) => r.arrears === 2).length],
                  ["Failed in 3 & above", a.results.filter((r) => r.arrears >= 3).length],
                ].map(([label, count]) => (
                  <div key={String(label)} className="flex items-center gap-3">
                    <span className="w-44 shrink-0 text-sm text-muted-foreground">{label}</span>
                    <div className="h-6 flex-1 overflow-hidden rounded-md bg-muted">
                      <div
                        className="h-full bg-primary/80"
                        style={{
                          width: `${a.totalStudents ? (Number(count) / a.totalStudents) * 100 : 0}%`,
                        }}
                      />
                    </div>
                    <span className="w-8 text-right font-mono text-sm">{count}</span>
                  </div>
                ))}
              </div>
              <dl className="mt-6 grid grid-cols-2 gap-4 border-t border-border pt-4 text-sm">
                <div>
                  <dt className="text-muted-foreground">Class average</dt>
                  <dd className="font-display text-xl">{fmt(a.classAverage)}%</dd>
                </div>
                <div>
                  <dt className="text-muted-foreground">Subjects</dt>
                  <dd className="font-display text-xl">{subjects.length}</dd>
                </div>
              </dl>
            </Panel>
          </div>
        )}

        {tab === "particulars" && (
          <Panel title="Particulars (category-wise)">
            <div className="overflow-x-auto">
              <table className="w-full min-w-[900px] border-collapse text-sm">
                <thead>
                  <tr className="border-b border-border text-left">
                    <Th>S.No.</Th>
                    <Th>Particulars</Th>
                    <Th align="right">Total</Th>
                    {CATEGORY_KEYS.map((k) => (
                      <Th key={k} align="right">
                        <span className="font-mono text-[11px]">{k}</span>
                      </Th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {a.particulars.map((p, i) => (
                    <tr key={p.label} className="border-b border-border/60 hover:bg-muted/50">
                      <Td>{i + 1}</Td>
                      <Td>{p.label}</Td>
                      <Td align="right" mono>
                        <strong>{p.total}</strong>
                      </Td>
                      {CATEGORY_KEYS.map((k) => (
                        <Td key={k} align="right" mono>
                          {p.byCategory[k]}
                        </Td>
                      ))}
                    </tr>
                  ))}
                  <tr className="bg-muted/60">
                    <Td>{a.particulars.length + 1}</Td>
                    <Td>Pass percentage</Td>
                    <Td align="right" mono>
                      <strong>{fmt(a.passPercent, 2)}%</strong>
                    </Td>
                    {CATEGORY_KEYS.map((k) => (
                      <Td key={k} />
                    ))}
                  </tr>
                </tbody>
              </table>
            </div>
            <ul className="mt-4 grid gap-1 text-xs text-muted-foreground sm:grid-cols-2">
              {CATEGORY_KEYS.map((k) => (
                <li key={k}>
                  <span className="font-mono text-accent">{k}</span> — {CATEGORY_LABEL[k]} (
                  {a.categoryCounts[k]})
                </li>
              ))}
            </ul>
          </Panel>
        )}

        {tab === "subjects" && (
          <Panel title="Subject wise performance">
            <div className="overflow-x-auto">
              <table className="w-full min-w-[820px] border-collapse text-sm">
                <thead>
                  <tr className="border-b border-border text-left">
                    <Th>S.No</Th>
                    <Th>Code</Th>
                    <Th>Subject</Th>
                    <Th>Staff</Th>
                    <Th align="right">Appeared</Th>
                    <Th align="right">Absent</Th>
                    <Th align="right">Passed</Th>
                    <Th align="right">Failed</Th>
                    <Th align="right">Pass %</Th>
                    <Th align="right">Avg</Th>
                    <Th align="right">High</Th>
                  </tr>
                </thead>
                <tbody>
                  {a.subjectStats.map((s, i) => (
                    <tr key={s.subject.code} className="border-b border-border/60 hover:bg-muted/50">
                      <Td>{i + 1}</Td>
                      <Td mono>{s.subject.code}</Td>
                      <Td>{s.subject.name}</Td>
                      <Td>{s.subject.staff || "—"}</Td>
                      <Td align="right" mono>{s.appeared}</Td>
                      <Td align="right" mono>{s.absent}</Td>
                      <Td align="right" mono>{s.passed}</Td>
                      <Td align="right" mono>{s.failed}</Td>
                      <Td align="right" mono>{fmt(s.passPercent)}</Td>
                      <Td align="right" mono>{fmt(s.average)}</Td>
                      <Td align="right" mono>{s.highest}</Td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Panel>
        )}

        {tab === "students" && (
          <Panel
            title={`Student results (${visibleResults.length})`}
            actions={
              <div className="flex flex-wrap items-center gap-2">
                <input
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  placeholder="Search name or reg no"
                  className="w-56 rounded-md border border-input bg-background px-3 py-1.5 text-sm"
                />
                {(["all", "pass", "fail", "absent"] as const).map((f) => (
                  <button
                    key={f}
                    onClick={() => setFilter(f)}
                    className={`rounded-full px-3 py-1 text-xs font-medium capitalize transition-colors ${
                      filter === f
                        ? "bg-primary text-primary-foreground"
                        : "bg-muted text-muted-foreground hover:text-foreground"
                    }`}
                  >
                    {f}
                  </button>
                ))}
              </div>
            }
          >
            <div className="overflow-x-auto">
              <table className="w-full min-w-[900px] border-collapse text-sm">
                <thead>
                  <tr className="border-b border-border text-left">
                    <Th align="right">Rank</Th>
                    <Th>Reg No</Th>
                    <Th>Name</Th>
                    {subjects.map((s) => (
                      <Th key={s.code} align="right">
                        <span className="font-mono text-[11px]">{s.code}</span>
                      </Th>
                    ))}
                    <Th align="right">Total</Th>
                    <Th align="right">%</Th>
                    <Th align="right">Arrears</Th>
                    <Th align="right">Absent</Th>
                    <Th>Status</Th>
                  </tr>
                </thead>
                <tbody>
                  {visibleResults.map((r) => (
                    <tr key={r.student.reg} className="border-b border-border/60 hover:bg-muted/50">
                      <Td align="right" mono>{r.rank}</Td>
                      <Td mono>{r.student.reg}</Td>
                      <Td>{r.student.name}</Td>
                      {r.student.marks.map((m, i) => (
                        <Td key={i} align="right" mono>
                          <span
                            className={
                              isAbsent(m)
                                ? "text-warning"
                                : typeof m === "number" && m < passMark
                                  ? "text-destructive"
                                  : ""
                            }
                          >
                            {m === null ? "—" : m}
                          </span>
                        </Td>
                      ))}
                      <Td align="right" mono>{r.total}</Td>
                      <Td align="right" mono>{fmt(r.percent)}</Td>
                      <Td align="right" mono>{r.arrears}</Td>
                      <Td align="right" mono>{r.absents}</Td>
                      <Td>
                        <span
                          className={`rounded-full px-2 py-0.5 text-xs font-semibold ${
                            r.status === "PASS"
                              ? "bg-success/15 text-success"
                              : "bg-destructive/15 text-destructive"
                          }`}
                        >
                          {r.status}
                        </span>
                      </Td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Panel>
        )}

        {tab === "toppers" && (
          <Panel title="Topper list">
            <ol className="grid gap-3 sm:grid-cols-2">
              {a.toppers.map((r) => (
                <li
                  key={r.student.reg}
                  className="flex items-center gap-4 rounded-lg border border-border bg-card p-4"
                >
                  <span className="font-display text-2xl text-accent">
                    {["I", "II", "III", "IV", "V"][r.rank - 1] ?? r.rank}
                  </span>
                  <div className="min-w-0">
                    <p className="truncate font-medium">{r.student.name}</p>
                    <p className="font-mono text-xs text-muted-foreground">{r.student.reg}</p>
                  </div>
                  <div className="ml-auto text-right">
                    <p className="font-display text-lg">{r.total}</p>
                    <p className="font-mono text-xs text-muted-foreground">{fmt(r.percent, 2)}%</p>
                  </div>
                </li>
              ))}
            </ol>
          </Panel>
        )}
      </main>

      <footer className="border-t border-border py-8 text-center text-xs text-muted-foreground">
        Prepared by · HOD · Vice Principal · Principal
      </footer>
    </div>
  );
}

function Stat({
  label,
  value,
  hint,
  tone = "neutral",
}: {
  label: string;
  value: string;
  hint?: string;
  tone?: "neutral" | "good" | "bad" | "warn";
}) {
  const toneClass =
    tone === "good"
      ? "text-success"
      : tone === "bad"
        ? "text-destructive"
        : tone === "warn"
          ? "text-warning"
          : "text-foreground";
  return (
    <div className="rounded-lg border border-border bg-card p-5">
      <p className="text-xs uppercase tracking-[0.18em] text-muted-foreground">{label}</p>
      <p className={`mt-2 font-display text-3xl ${toneClass}`}>{value}</p>
      {hint && <p className="mt-1 text-xs text-muted-foreground">{hint}</p>}
    </div>
  );
}

function Panel({
  title,
  actions,
  children,
}: {
  title: string;
  actions?: React.ReactNode;
  children: React.ReactNode;
}) {
  return (
    <section className="rounded-lg border border-border bg-card p-6">
      <div className="mb-5 flex flex-wrap items-center justify-between gap-3">
        <h2 className="font-display text-lg tracking-tight">{title}</h2>
        {actions}
      </div>
      {children}
    </section>
  );
}

function Th({
  children,
  align = "left",
}: {
  children?: React.ReactNode;
  align?: "left" | "right";
}) {
  return (
    <th
      className={`px-3 py-2 text-xs font-semibold uppercase tracking-wider text-muted-foreground ${
        align === "right" ? "text-right" : "text-left"
      }`}
    >
      {children}
    </th>
  );
}

function Td({
  children,
  align = "left",
  mono = false,
}: {
  children?: React.ReactNode;
  align?: "left" | "right";
  mono?: boolean;
}) {
  return (
    <td
      className={`px-3 py-2 ${align === "right" ? "text-right" : ""} ${
        mono ? "font-mono text-xs" : ""
      }`}
    >
      {children}
    </td>
  );
}
