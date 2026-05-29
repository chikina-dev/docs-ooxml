import { benchmarkDocxWriters } from "../src/ooxml/docx.ts";
import { createPipelineFromLexicalJson } from "../src/pipeline/createPipeline.ts";

type Scenario = {
  name: string;
  paragraphs: number;
  listItems: number;
  iterations: number;
};

const SCENARIOS: Scenario[] = [
  { name: "small", paragraphs: 20, listItems: 10, iterations: 200 },
  { name: "medium", paragraphs: 200, listItems: 50, iterations: 60 },
  { name: "large", paragraphs: 1000, listItems: 120, iterations: 15 },
];

const metadata = {
  title: "GitHub Actions DOCX Writer Benchmark",
  creator: "Docs OOXML CI",
  createdAt: new Date("2026-05-29T00:00:00.000Z"),
};

const rows = SCENARIOS.flatMap((scenario) => {
  const projection = createPipelineFromLexicalJson(createLexicalFixture(scenario)).outputProjection;
  benchmarkDocxWriters(projection, metadata, 3);

  return benchmarkDocxWriters(projection, metadata, scenario.iterations).map((result) => ({
    scenario: scenario.name,
    paragraphs: scenario.paragraphs,
    listItems: scenario.listItems,
    strategy: result.strategy,
    iterations: result.iterations,
    totalMs: result.durationMs,
    avgMs: result.durationMs / result.iterations,
    sizeBytes: result.sizeBytes,
  }));
});

const table = [
  "| Scenario | Paragraphs | List Items | Strategy | Iterations | Total ms | Avg ms/run | DOCX size |",
  "| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: |",
  ...rows.map(
    (row) =>
      `| ${row.scenario} | ${row.paragraphs} | ${row.listItems} | ${row.strategy} | ${row.iterations} | ${row.totalMs.toFixed(2)} | ${row.avgMs.toFixed(4)} | ${formatBytes(row.sizeBytes)} |`,
  ),
].join("\n");

const summary = `## DOCX Writer Benchmark

Measured in GitHub Actions with Bun on \`ubuntu-latest\`.

${table}
`;

console.log(summary);

const bun = globalThis as typeof globalThis & {
  Bun?: {
    env: Record<string, string | undefined>;
    write(path: string, content: string): Promise<number>;
  };
};

const stepSummaryPath = bun.Bun?.env.GITHUB_STEP_SUMMARY;
if (stepSummaryPath) {
  await bun.Bun?.write(stepSummaryPath, `${summary}\n`);
}

function createLexicalFixture(scenario: Scenario) {
  return {
    root: {
      children: [
        {
          type: "heading",
          tag: "h1",
          children: [{ type: "text", text: `${scenario.name} benchmark`, format: 1 }],
        },
        ...Array.from({ length: scenario.paragraphs }, (_, index) => ({
          type: "paragraph",
          children: [
            { type: "text", text: `Paragraph ${index + 1}: `, format: 1 },
            {
              type: "text",
              text: "This paragraph exercises XML escaping, run projection, and ZIP packaging. ",
              format: 0,
            },
            { type: "text", text: "A&B <quoted> content.", format: index % 2 === 0 ? 10 : 0 },
          ],
        })),
        {
          type: "list",
          listType: "number",
          children: Array.from({ length: scenario.listItems }, (_, index) => ({
            type: "listitem",
            children: [{ type: "text", text: `Measured list item ${index + 1}`, format: 0 }],
          })),
        },
      ],
    },
  };
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  return `${(bytes / 1024).toFixed(1)} KB`;
}
