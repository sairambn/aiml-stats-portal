import { useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import {
  analyse,
  fmt,
  isAbsent,
  type Student,
  type Subject,
  CATEGORY_KEYS,
} from "@/lib/analysis";
import { meta as seedMeta } from "@/data/seed";
import { exportAnalysisSheet, type ExportMeta } from "@/lib/export-sheet";
import { parseWorkbook } from "@/lib/parse-sheet";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Result Analysis Portal · Exam Cell" },
      {
        name: "description",
        content:
          "Professional internal assessment result analysis for exam cell — particulars, subject-wise performance, toppers and arrear groups.",
      },
    ],
  }),
  component: Portal,
});

type TabKey = "overview" | "particulars" | "subjects" | "students" | "toppers";

const TABS: { key: TabKey; label: string }[] = [
  { key: "overview", label: "Overview" },
  { key: "particulars", label: "Particulars" },
  { key: "subjects", label: "Subjects" },
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
  const [students, setStudents] = useState<Student[]>([]);
  const [subjects, setSubjects] = useState<Subject[]>([]);
  const [meta, setMeta] = useState<ExportMeta>({
    institution: seedMeta.institution,
    department: seedMeta.department,
    title: seedMeta.title,
    year: seedMeta.year,
    semester: seedMeta.semester,
    batch: seedMeta.batch,
    section: seedMeta.section,
  });
  const [passMark, setPassMark] = useState(seedMeta.passMark);
  const [tab, setTab] = useState<TabKey>("overview");
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<"all" | "pass" | "fail" | "absent">("all");
  const [notice, setNotice] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const fileRef = useRef<HTMLInputElement>(null);

  const a = useMemo(
    () => analyse(students, subjects, passMark),
    [students, subjects, passMark],
  );

  const rankedStudents = useMemo(
    () =>
      [...a.results].sort(
        (x, y) => x.rank - y.rank || x.student.name.localeCompare(y.student.name),
      ),
    [a.results],
  );

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
        const buf = e.target?.result;
        if (!(buf instanceof ArrayBuffer)) throw new Error("Could not read file");
        const parsed = parseWorkbook(buf);
        setStudents(parsed.students);
        setSubjects(parsed.subjects);
        setMeta((prev) => ({
          ...prev,
          ...parsed.meta,
          institution: parsed.meta.institution || prev.institution,
          department: parsed.meta.department || prev.department,
          title: parsed.meta.title || prev.title,
          year: parsed.meta.year || prev.year,
          semester: parsed.meta.semester || prev.semester,
          batch: parsed.meta.batch || prev.batch,
          section: parsed.meta.section || prev.section,
        }));
        const codes = parsed.subjects.map((s) => s.code).join(", ");
        setNotice(
          `Loaded ${parsed.students.length} students · ${parsed.subjects.length} subjects (${codes}) from ${file.name}`,
        );
        setTab("overview");
      } catch (err) {
        setNotice(err instanceof Error ? err.message : "Could not read that file");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function clearData() {
    setStudents([]);
    setSubjects([]);
    setNotice(null);
    setQuery("");
    setFilter("all");
    setTab("overview");
  }

  async function exportSheet() {
    try {
      setExporting(true);
      setNotice(null);
      await exportAnalysisSheet(a, subjects, meta, passMark);
      setNotice("Analysis Excel downloaded — ready for exam cell");
    } catch (err) {
      setNotice(err instanceof Error ? err.message : "Export failed");
    } finally {
      setExporting(false);
    }
  }

  return (
    <div className="flex min-h-screen bg-background text-foreground">
      <aside
        className={`print:hidden sticky top-0 flex h-screen shrink-0 flex-col border-r border-border bg-card transition-all duration-200 ${
          sidebarOpen ? "w-80" : "w-0 overflow-hidden border-0"
        }`}
      >
        <div className="flex h-full w-80 flex-col">
          <div className="border-b border-border px-5 py-5">
            <p className="font-mono text-[10px] uppercase tracking-[0.22em] text-accent">
              Exam Cell
            </p>
            <h2 className="mt-1 font-display text-lg leading-tight tracking-tight">
              Result Analysis
            </h2>
            <p className="mt-1 text-xs text-muted-foreground line-clamp-2">
              {meta.department}
            </p>
          </div>

          <div className="border-b border-border px-5 py-4 space-y-2.5 text-xs">
            <label className="block">
              <span className="text-muted-foreground">Assessment</span>
              <input
                className="mt-0.5 w-full rounded border border-input bg-background px-2 py-1 text-xs"
                value={meta.title}
                onChange={(e) => setMeta((m) => ({ ...m, title: e.target.value }))}
              />
            </label>
            <label className="block">
              <span className="text-muted-foreground">Department</span>
              <input
                className="mt-0.5 w-full rounded border border-input bg-background px-2 py-1 text-xs"
                value={meta.department}
                onChange={(e) => setMeta((m) => ({ ...m, department: e.target.value }))}
              />
            </label>
            <div className="grid grid-cols-2 gap-2">
              <label className="block">
                <span className="text-muted-foreground">Year</span>
                <input
                  className="mt-0.5 w-full rounded border border-input bg-background px-2 py-1 text-xs"
                  value={meta.year}
                  onChange={(e) => setMeta((m) => ({ ...m, year: e.target.value }))}
                />
              </label>
              <label className="block">
                <span className="text-muted-foreground">Semester</span>
                <input
                  className="mt-0.5 w-full rounded border border-input bg-background px-2 py-1 text-xs"
                  value={meta.semester}
                  onChange={(e) => setMeta((m) => ({ ...m, semester: e.target.value }))}
                />
              </label>
              <label className="block">
                <span className="text-muted-foreground">Batch</span>
                <input
                  className="mt-0.5 w-full rounded border border-input bg-background px-2 py-1 text-xs"
                  value={meta.batch}
                  onChange={(e) => setMeta((m) => ({ ...m, batch: e.target.value }))}
                />
              </label>
              <label className="block">
                <span className="text-muted-foreground">Section</span>
                <input
                  className="mt-0.5 w-full rounded border border-input bg-background px-2 py-1 text-xs"
                  value={meta.section}
                  onChange={(e) => setMeta((m) => ({ ...m, section: e.target.value }))}
                />
              </label>
            </div>
            <div className="flex items-center justify-between pt-1">
              <span className="text-muted-foreground">Pass mark</span>
              <input
                type="number"
                min={1}
                max={100}
                value={passMark}
                onChange={(e) =>
                  setPassMark(Math.max(1, Math.min(100, Number(e.target.value) || 1)))
                }
                className="w-14 rounded border border-input bg-background px-2 py-0.5 text-right font-mono text-xs"
              />
            </div>
          </div>

          <div className="flex-1 overflow-y-auto px-5 py-4">
            {students.length === 0 ? (
              <p className="text-xs text-muted-foreground">
                Upload a mark sheet to see subjects and student names here.
              </p>
            ) : (
              <>
                <p className="mb-3 font-mono text-[10px] uppercase tracking-[0.18em] text-muted-foreground">
                  Subjects · Full forms ({subjects.length})
                </p>
                <ul className="space-y-3">
                  {subjects.map((s) => {
                    const st = a.subjectStats.find((x) => x.subject.code === s.code);
                    return (
                      <li
                        key={s.code}
                        className="rounded-md border border-border/70 bg-muted/20 p-3"
                      >
                        <div className="flex items-baseline justify-between gap-2">
                          <span className="font-mono text-[11px] font-semibold text-accent">
                            {s.code}
                          </span>
                          {st && (
                            <span className="font-mono text-[10px] text-muted-foreground">
                              {fmt(st.passPercent)}%
                            </span>
                          )}
                        </div>
                        <p className="mt-1 text-xs font-medium leading-snug">{s.name}</p>
                        {s.staff ? (
                          <p className="mt-1 text-[11px] text-muted-foreground">{s.staff}</p>
                        ) : null}
                      </li>
                    );
                  })}
                </ul>

                <p className="mb-3 mt-6 font-mono text-[10px] uppercase tracking-[0.18em] text-muted-foreground">
                  Students · Ranked ({rankedStudents.length})
                </p>
                <ul className="space-y-1 pb-6">
                  {rankedStudents.slice(0, 40).map((r) => (
                    <li
                      key={r.student.reg}
                      className="flex items-center gap-2 rounded px-1.5 py-1 text-[11px] hover:bg-muted/40"
                    >
                      <span className="w-5 shrink-0 font-mono text-muted-foreground">
                        {r.rank}
                      </span>
                      <span className="min-w-0 flex-1 truncate font-medium">
                        {r.student.name}
                      </span>
                      <span
                        className={`shrink-0 rounded px-1.5 py-0.5 font-mono text-[10px] ${
                          r.status === "PASS"
                            ? "bg-success/15 text-success"
                            : "bg-destructive/15 text-destructive"
                        }`}
                      >
                        {r.status === "PASS" ? "P" : "F"}
                      </span>
                    </li>
                  ))}
                  {rankedStudents.length > 40 && (
                    <li className="px-1.5 text-[11px] text-muted-foreground">
                      +{rankedStudents.length - 40} more
                    </li>
                  )}
                </ul>
              </>
            )}
          </div>
        </div>
      </aside>

      <div className="flex min-w-0 flex-1 flex-col">
        <header className="border-b border-border bg-card">
          <div className="flex flex-wrap items-start justify-between gap-4 px-6 py-5">
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  className="print:hidden rounded-md border border-border px-2 py-1 text-xs text-muted-foreground hover:bg-muted"
                  onClick={() => setSidebarOpen((v) => !v)}
                  title="Toggle sidebar"
                >
                  {sidebarOpen ? "Hide panel" : "Show panel"}
                </button>
                <p className="font-mono text-[10px] uppercase tracking-[0.22em] text-accent">
                  {meta.institution}
                </p>
              </div>
              <h1 className="mt-2 font-display text-2xl tracking-tight sm:text-3xl">
                {meta.title || "Result Analysis"}
              </h1>
              <p className="mt-1 text-sm text-muted-foreground">
                Year {meta.year} · Sem {meta.semester} · Batch {meta.batch} · Sec{" "}
                {meta.section}
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2 print:hidden">
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
              <button
                className="btn-outline"
                disabled={exporting || a.totalStudents === 0}
                onClick={() => void exportSheet()}
              >
                {exporting ? "Preparing…" : "Export analysis"}
              </button>
              <button className="btn-ghost" onClick={() => window.print()}>
                Print
              </button>
              {students.length > 0 && (
                <button className="btn-ghost" onClick={clearData}>
                  Clear
                </button>
              )}
            </div>
          </div>

          {notice && (
            <p className="border-t border-accent/30 bg-accent/10 px-6 py-2 text-sm text-accent-foreground">
              {notice}
            </p>
          )}

          {students.length > 0 && (
            <nav className="print:hidden flex gap-1 overflow-x-auto border-t border-border px-6">
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
            </nav>
          )}
        </header>

        <main className="mx-auto w-full max-w-6xl flex-1 space-y-6 px-6 py-8">
          {students.length === 0 ? (
            <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-border bg-card px-6 py-20 text-center">
              <p className="font-mono text-[10px] uppercase tracking-[0.22em] text-accent">
                Exam Cell
              </p>
              <h2 className="mt-3 font-display text-2xl tracking-tight">
                Upload a mark sheet to begin
              </h2>
              <p className="mt-2 max-w-md text-sm text-muted-foreground">
                No data is stored. On refresh this page is empty until you upload
                an Excel mark sheet. Analysis and export stay on your device.
              </p>
              <button
                className="btn-primary mt-8"
                onClick={() => fileRef.current?.click()}
              >
                Upload mark sheet
              </button>
              <p className="mt-4 text-xs text-muted-foreground">
                Accepts .xlsx · MARK sheet with REG. NO. and subject codes
              </p>
            </div>
          ) : (
            <>
              <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                <Stat
                  label="Total students"
                  value={String(a.totalStudents)}
                  hint={`${a.appeared} appeared`}
                />
                <Stat
                  label="Passed all subjects"
                  value={String(a.passedAll)}
                  hint={`${fmt(a.passPercent)}% pass`}
                  tone="good"
                />
                <Stat
                  label="Failed"
                  value={String(a.failed)}
                  hint={`${a.results.filter((r) => r.arrears >= 3).length} with 3+ arrears`}
                  tone="bad"
                />
                <Stat
                  label="Absentees"
                  value={String(a.totalAbsentEntries)}
                  hint={`${a.studentsWithAbsence} students affected`}
                  tone="warn"
                />
              </section>

              {tab === "overview" && (
                <div className="grid gap-6 lg:grid-cols-2">
                  <Panel title="Pass percentage per subject">
                    <div className="space-y-4">
                      {a.subjectStats.map((s) => (
                        <div key={s.subject.code}>
                          <div className="flex items-baseline justify-between gap-3 text-sm">
                            <span className="font-medium">
                              <span className="font-mono text-xs text-accent">
                                {s.subject.code}
                              </span>{" "}
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
                      {(
                        [
                          ["No arrears (all clear)", a.results.filter((r) => r.arrears === 0).length],
                          ["Failed in 1 subject", a.results.filter((r) => r.arrears === 1).length],
                          ["Failed in 2 subjects", a.results.filter((r) => r.arrears === 2).length],
                          ["Failed in 3 & above", a.results.filter((r) => r.arrears >= 3).length],
                        ] as const
                      ).map(([label, count]) => (
                        <div key={label} className="flex items-center gap-3">
                          <span className="w-44 shrink-0 text-sm text-muted-foreground">{label}</span>
                          <div className="h-6 flex-1 overflow-hidden rounded-md bg-muted">
                            <div
                              className="h-full bg-primary/80"
                              style={{
                                width: `${a.totalStudents ? (count / a.totalStudents) * 100 : 0}%`,
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
            </>
          )}
        </main>

        <footer className="border-t border-border py-6 text-center text-xs text-muted-foreground">
          Prepared by · HOD · Vice Principal · Principal
        </footer>
      </div>
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
