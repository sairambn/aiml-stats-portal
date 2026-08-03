import { createFileRoute } from "@tanstack/react-router";
import { Portal } from "./portal";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Result Analysis Portal · Exam Cell" },
      {
        name: "description",
        content:
          "Professional internal assessment result analysis for exam cell — diagrams, subject-wise performance, toppers and arrear groups.",
      },
    ],
  }),
  component: Portal,
});
