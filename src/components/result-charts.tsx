import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import type { Analysis } from "@/lib/analysis";
import { fmt } from "@/lib/analysis";

const COLORS = {
  accent: "#3b82f6",
  good: "#16a34a",
  bad: "#dc2626",
  warn: "#d97706",
  muted: "#94a3b8",
  pie: ["#16a34a", "#3b82f6", "#d97706", "#dc2626"],
};

export function SubjectPassChart({ analysis }: { analysis: Analysis }) {
  const data = analysis.subjectStats.map((s) => ({
    code: s.subject.code,
    name: s.subject.name,
    pass: Number(fmt(s.passPercent, 1)),
    fail: Math.max(0, Number(fmt(100 - s.passPercent, 1))),
  }));

  return (
    <div className="h-72 w-full">
      <ResponsiveContainer width="100%" height="100%">
        <BarChart data={data} margin={{ top: 8, right: 8, left: 0, bottom: 48 }}>
          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
          <XAxis
            dataKey="code"
            tick={{ fontSize: 11 }}
            interval={0}
            angle={-25}
            textAnchor="end"
            height={50}
          />
          <YAxis domain={[0, 100]} tick={{ fontSize: 11 }} unit="%" width={40} />
          <Tooltip
            formatter={(value: number, key: string) => [
              `${value}%`,
              key === "pass" ? "Pass %" : "Fail %",
            ]}
            labelFormatter={(_, payload) => {
              const p = payload?.[0]?.payload as
                | { code: string; name: string }
                | undefined;
              return p ? `${p.code} — ${p.name}` : "";
            }}
          />
          <Legend />
          <Bar dataKey="pass" name="Pass %" stackId="a" fill={COLORS.good} />
          <Bar
            dataKey="fail"
            name="Fail %"
            stackId="a"
            fill={COLORS.bad}
            radius={[4, 4, 0, 0]}
          />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

export function ArrearPieChart({ analysis }: { analysis: Analysis }) {
  const data = [
    {
      name: "All clear",
      value: analysis.results.filter((r) => r.arrears === 0).length,
    },
    {
      name: "1 arrear",
      value: analysis.results.filter((r) => r.arrears === 1).length,
    },
    {
      name: "2 arrears",
      value: analysis.results.filter((r) => r.arrears === 2).length,
    },
    {
      name: "3+ arrears",
      value: analysis.results.filter((r) => r.arrears >= 3).length,
    },
  ].filter((d) => d.value > 0);

  if (!data.length) {
    return (
      <p className="py-10 text-center text-sm text-muted-foreground">No data</p>
    );
  }

  return (
    <div className="h-72 w-full">
      <ResponsiveContainer width="100%" height="100%">
        <PieChart>
          <Pie
            data={data}
            dataKey="value"
            nameKey="name"
            cx="50%"
            cy="50%"
            innerRadius={55}
            outerRadius={95}
            paddingAngle={2}
            label={({ name, percent }) =>
              `${name} ${(percent * 100).toFixed(0)}%`
            }
          >
            {data.map((_, i) => (
              <Cell key={i} fill={COLORS.pie[i % COLORS.pie.length]} />
            ))}
          </Pie>
          <Tooltip formatter={(v: number) => [v, "Students"]} />
          <Legend />
        </PieChart>
      </ResponsiveContainer>
    </div>
  );
}

export function PassFailChart({ analysis }: { analysis: Analysis }) {
  const data = [
    { name: "Passed all", value: analysis.passedAll, fill: COLORS.good },
    { name: "Failed", value: analysis.failed, fill: COLORS.bad },
  ];

  return (
    <div className="h-56 w-full">
      <ResponsiveContainer width="100%" height="100%">
        <BarChart data={data} layout="vertical" margin={{ left: 16, right: 24 }}>
          <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e2e8f0" />
          <XAxis type="number" tick={{ fontSize: 11 }} allowDecimals={false} />
          <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={90} />
          <Tooltip />
          <Bar dataKey="value" name="Students" radius={[0, 6, 6, 0]}>
            {data.map((d, i) => (
              <Cell key={i} fill={d.fill} />
            ))}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
