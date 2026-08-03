import { useMemo, useRef, useState } from "react";
import {
  analyse,
  fmt,
  type Student,
  type Subject,
} from "@/lib/analysis";
import {
  meta as seedMeta,
  seedStudents,
  seedSubjects,
} from "@/data/seed";
import { exportAnalysisSheet, type ExportMeta } from "@/lib/export-sheet";
import { parseWorkbook } from "@/lib/parse-sheet";
import {
  ArrearPieChart,
  PassFailChart,
  SubjectPassChart,
} from "@/components/result-charts";

type TabKey = "overview" | "subjects" | "students" | "toppers";

const TABS: { key: TabKey; label: string }[] = [
  { key: "overview", label: "Overview" },
  { key: "subjects", label: "Subjects" },
  { key: "students", label: "Students" },
  { key: "toppers", label: "Toppers" },
];

export function Portal() {
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
  const [notice, setNotice] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const a = useMemo(
    () => analyse(students, subjects, passMark),
    [students, subjects, passMark],
  );

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return a.results;
    return a.results.filter(
      (r) =>
        r.student.name.toLowerCase().includes(q) ||
        r.student.reg.includes(q),
    );
  }, [a.results, query]);

  function handleUpload(file: File) {
    setUploading(true);
    setNotice(null);
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
        setNotice(
          `Loaded ${parsed.students.length} students · ${parsed.subjects.length} subjects from ${file.name}`,
        );
        setTab("overview");
      } catch (err) {
        setNotice(err instanceof Error ? err.message : "Could not read that file");
      } finally {
        setUploading(false);
      }
    };
    reader.onerror = () => {
      setNotice("Could not read that file");
      setUploading(false);
    };
    reader.readAsArrayBuffer(file);
  }

  function loadDemo() {
    setStudents(seedStudents.map((s) => ({ ...s, marks: [...s.marks] })));
    setSubjects(seedSubjects.map((s) => ({ ...s })));
    setMeta({
      institution: seedMeta.institution,
      department: seedMeta.department,
      title: seedMeta.title,
      year: seedMeta.year,
      semester: seedMeta.semester,
      batch: seedMeta.batch,
      section: seedMeta.section,
    });
    setPassMark(seedMeta.passMark);
    setNotice(
      `Demo loaded · ${seedStudents.length} students · ${seedSubjects.length} subjects`,
    );
    setTab("overview");
  }

  function clearData() {
    setStudents([]);
    setSubjects([]);
    setNotice(null);
    setQuery("");
    setTab("overview");
  }

  async function exportSheet() {
    try {
      setExporting(true);
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
        className={`print:hidden sticky top-0 flex h-screen shrink-0 flex-col border-r border-border bg-card transition-all ${
          sidebarOpen ? "w-72" : "w-0 overflow-hidden border-0"
        }`}
      >
        <div className="flex h-full w-72 flex-col">
          <div className="border-b border-border px-5 py-5">
            <p className="font-mono text-[10px] uppercase tracking-[0.22em] text-accent">
              Exam Cell
            </p>
            <h2 className="mt-1 font-display text-lg tracking-tight">Result Analysis</h2>
            <p className="mt-1 text-xs text-muted-foreground line-clamp-2">{meta.department}</p>
          </div>
          <div className="space-y-2 border-b border-border px-5 py-4 text-xs">
            <label className="block">
              <span className="text-muted-foreground">Assessment</span>
              <input
                className="mt-0.5 w-full rounded border border-input bg-background px-2 py-1 text-xs"
                value={meta.title}
                onChange={(e) => setMeta((m) => ({ ...m, title: e.target.value }))}
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
            {subjects.length === 0 ? (
              <p className="text-xs text-muted-foreground">Upload or load demo to begin.</p>
            ) : (
              <ul className="space-y-2">
                {subjects.map((s) => (
                  <li key={s.code} className="rounded-md border border-border/70 p-2 text-xs">
                    <span className="font-mono font-semibold text-accent">{s.code}</span>
                    <p className="mt-0.5 font-medium">{s.name}</p>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </div>
      </aside>

      <div className="flex min-w-0 flex-1 flex-col">
        <header className="border-b border-border bg-card">
          <div className="flex flex-wrap items-start justify-between gap-4 px-6 py-5">
            <div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  className="print:hidden rounded-md border border-border px-2 py-1 text-xs text-muted-foreground hover:bg-muted"
                  onClick={() => setSidebarOpen((v) => !v)}
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
              <button
                className="btn-primary"
                disabled={uploading}
                onClick={() => fileRef.current?.click()}
              >
                {uploading ? "Uploading…" : "Upload mark sheet"}
              </button>
              <button
                className="btn-outline"
                disabled={exporting || a.totalStudents === 0}
                onClick={() => void exportSheet()}
              >
                {exporting ? "Preparing…" : "Export analysis"}
              </button>
              {students.length > 0 && (
                <button className="btn-ghost" onClick={clearData}>
                  Clear
                </button>
              )}
            </div>
          </div>
          {notice && (
            <p className="border-t border-accent/30 bg-accent/10 px-6 py-2 text-sm">
              {notice}
            </p>
          )}
          {students.length > 0 && (
            <nav className="print:hidden flex gap-1 border-t border-border px-6">
              {TABS.map((t) => (
                <button
                  key={t.key}
                  onClick={() => setTab(t.key)}
                  className={`whitespace-nowrap border-b-2 px-4 py-3 text-sm font-medium ${
                    tab === t.key
                      ? "border-accent text-foreground"
                      : "border-transparent text-muted-foreground"
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
            <div
              className={`flex flex-col items-center justify-center rounded-xl border-2 border-dashed px-6 py-20 text-center ${
                dragOver ? "border-accent bg-accent/10" : "border-border bg-card"
              }`}
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragOver(false);
                const f = e.dataTransfer.files?.[0];
                if (f) handleUpload(f);
              }}
            >
              <p className="font-mono text-[10px] uppercase tracking-[0.22em] text-accent">
                Exam Cell
              </p>
              <h2 className="mt-3 font-display text-2xl tracking-tight">
                {uploading ? "Reading mark sheet…" : "Upload a mark sheet to begin"}
              </h2>
              <p className="mt-2 max-w-md text-sm text-muted-foreground">
                Drag & drop an Excel file, or use demo data. Nothing is stored on the
                server.
              </p>
              <div className="mt-8 flex flex-wrap justify-center gap-3">
                <button
                  className="btn-primary"
                  disabled={uploading}
                  onClick={() => fileRef.current?.click()}
                >
                  {uploading ? "Uploading…" : "Upload mark sheet"}
                </button>
                <button className="btn-outline" disabled={uploading} onClick={loadDemo}>
                  Load demo data
                </button>
              </div>
            </div>
          ) : (
            <>
              <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                <Card label="Total students" value={String(a.totalStudents)} />
                <Card
                  label="Passed all"
                  value={String(a.passedAll)}
                  hint={`${fmt(a.passPercent)}%`}
                  good
                />
                <Card label="Failed" value={String(a.failed)} bad />
                <Card label="Absent entries" value={String(a.totalAbsentEntries)} />
              </section>

              {tab === "overview" && (
                <div className="grid gap-6 lg:grid-cols-2">
                  <Section title="Subject-wise pass">
                    <SubjectPassChart analysis={a} />
                  </Section>
                  <Section title="Arrear distribution">
                    <ArrearPieChart analysis={a} />
                  </Section>
                  <Section title="Pass vs fail">
                    <PassFailChart analysis={a} />
                  </Section>
                </div>
              )}

              {tab === "subjects" && (
                <Section title="Subject performance">
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="border-b text-left text-xs text-muted-foreground">
                          <th className="px-2 py-2">Code</th>
                          <th className="px-2 py-2">Subject</th>
                          <th className="px-2 py-2 text-right">Pass %</th>
                          <th className="px-2 py-2 text-right">Avg</th>
                        </tr>
                      </thead>
                      <tbody>
                        {a.subjectStats.map((s) => (
                          <tr key={s.subject.code} className="border-b border-border/50">
                            <td className="px-2 py-2 font-mono text-xs">{s.subject.code}</td>
                            <td className="px-2 py-2">{s.subject.name}</td>
                            <td className="px-2 py-2 text-right font-mono text-xs">
                              {fmt(s.passPercent)}
                            </td>
                            <td className="px-2 py-2 text-right font-mono text-xs">
                              {fmt(s.average)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </Section>
              )}

              {tab === "students" && (
                <Section title={`Students (${visible.length})`}>
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Search name or reg no"
                    className="mb-4 w-56 rounded-md border border-input bg-background px-3 py-1.5 text-sm"
                  />
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="border-b text-left text-xs text-muted-foreground">
                          <th className="px-2 py-2 text-right">Rank</th>
                          <th className="px-2 py-2">Reg</th>
                          <th className="px-2 py-2">Name</th>
                          <th className="px-2 py-2 text-right">Total</th>
                          <th className="px-2 py-2 text-right">%</th>
                          <th className="px-2 py-2">Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visible.map((r) => (
                          <tr key={r.student.reg} className="border-b border-border/50">
                            <td className="px-2 py-2 text-right font-mono text-xs">{r.rank}</td>
                            <td className="px-2 py-2 font-mono text-xs">{r.student.reg}</td>
                            <td className="px-2 py-2">{r.student.name}</td>
                            <td className="px-2 py-2 text-right font-mono text-xs">{r.total}</td>
                            <td className="px-2 py-2 text-right font-mono text-xs">
                              {fmt(r.percent)}
                            </td>
                            <td className="px-2 py-2">
                              <span
                                className={`rounded-full px-2 py-0.5 text-xs font-semibold ${
                                  r.status === "PASS"
                                    ? "bg-success/15 text-success"
                                    : "bg-destructive/15 text-destructive"
                                }`}
                              >
                                {r.status}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </Section>
              )}

              {tab === "toppers" && (
                <Section title="Toppers">
                  <ol className="grid gap-3 sm:grid-cols-2">
                    {a.toppers.map((r) => (
                      <li
                        key={r.student.reg}
                        className="flex items-center gap-4 rounded-lg border border-border p-4"
                      >
                        <span className="font-display text-2xl text-accent">
                          {["I", "II", "III", "IV", "V"][r.rank - 1] ?? r.rank}
                        </span>
                        <div className="min-w-0">
                          <p className="truncate font-medium">{r.student.name}</p>
                          <p className="font-mono text-xs text-muted-foreground">
                            {r.student.reg}
                          </p>
                        </div>
                        <div className="ml-auto text-right">
                          <p className="font-display text-lg">{r.total}</p>
                          <p className="font-mono text-xs text-muted-foreground">
                            {fmt(r.percent, 2)}%
                          </p>
                        </div>
                      </li>
                    ))}
                  </ol>
                </Section>
              )}
            </>
          )}
        </main>

        <footer className="border-t border-border py-6 text-center text-xs text-muted-foreground">
          Designed and crafted by{" "}
          <a
            href="https://bnsairam.vercel.app/"
            target="_blank"
            rel="noopener noreferrer"
            className="text-accent hover:underline"
          >
            https://bnsairam.vercel.app/
          </a>
        </footer>
      </div>
    </div>
  );
}

function Card({
  label,
  value,
  hint,
  good,
  bad,
}: {
  label: string;
  value: string;
  hint?: string;
  good?: boolean;
  bad?: boolean;
}) {
  return (
    <div className="rounded-lg border border-border bg-card p-5">
      <p className="text-xs uppercase tracking-[0.18em] text-muted-foreground">{label}</p>
      <p
        className={`mt-2 font-display text-3xl ${
          good ? "text-success" : bad ? "text-destructive" : ""
        }`}
      >
        {value}
      </p>
      {hint && <p className="mt-1 text-xs text-muted-foreground">{hint}</p>}
    </div>
  );
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section className="rounded-lg border border-border bg-card p-6">
      <h2 className="mb-5 font-display text-lg tracking-tight">{title}</h2>
      {children}
    </section>
  );
}
