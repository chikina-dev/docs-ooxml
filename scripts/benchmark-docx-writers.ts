import { createDocxBlob, type DocxWriteStrategy } from "../src/ooxml/docx.ts";
import { createPipelineFromLexicalJson } from "../src/pipeline/createPipeline.ts";

type Scenario = {
  name: string;
  paragraphs: number;
  listItems: number;
  iterations: number;
};

type Sample = {
  durationMs: number;
  sizeBytes: number;
};

type Row = {
  scenario: string;
  paragraphs: number;
  listItems: number;
  strategy: DocxWriteStrategy;
  iterations: number;
  medianAvgMs: number;
  minAvgMs: number;
  maxAvgMs: number;
  sizeBytes: number;
};

const SCENARIOS: Scenario[] = [
  { name: "small", paragraphs: 20, listItems: 10, iterations: 200 },
  { name: "medium", paragraphs: 200, listItems: 50, iterations: 60 },
  { name: "large", paragraphs: 1000, listItems: 120, iterations: 15 },
];

const SAMPLE_COUNT = 7;
const STRATEGIES: DocxWriteStrategy[] = ["naive", "optimized"];

const metadata = {
  title: "GitHub Actions DOCX Writer Benchmark",
  creator: "Docs OOXML CI",
  createdAt: new Date("2026-05-29T00:00:00.000Z"),
};

const rows = SCENARIOS.flatMap((scenario) => {
  const projection = createPipelineFromLexicalJson(createLexicalFixture(scenario)).outputProjection;
  const samples = new Map<DocxWriteStrategy, Sample[]>(
    STRATEGIES.map((strategy) => [strategy, []]),
  );

  for (const strategy of STRATEGIES) {
    runStrategy(strategy, projection, scenario.iterations, 3);
  }

  for (let sampleIndex = 0; sampleIndex < SAMPLE_COUNT; sampleIndex += 1) {
    const order = sampleIndex % 2 === 0 ? STRATEGIES : [...STRATEGIES].reverse();
    for (const strategy of order) {
      samples.get(strategy)?.push(runStrategy(strategy, projection, scenario.iterations));
    }
  }

  return STRATEGIES.map((strategy): Row => {
    const strategySamples = samples.get(strategy) ?? [];
    const averages = strategySamples.map((sample) => sample.durationMs / scenario.iterations);
    return {
      scenario: scenario.name,
      paragraphs: scenario.paragraphs,
      listItems: scenario.listItems,
      strategy,
      iterations: scenario.iterations,
      medianAvgMs: median(averages),
      minAvgMs: Math.min(...averages),
      maxAvgMs: Math.max(...averages),
      sizeBytes: strategySamples.at(-1)?.sizeBytes ?? 0,
    };
  });
});

const table = [
  "| Scenario | Paragraphs | List Items | Strategy | Iterations/sample | Median avg ms/run | Min-Max avg | vs naive | DOCX size |",
  "| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: |",
  ...rows.map((row) => {
    const baseline = rows.find(
      (item) => item.scenario === row.scenario && item.strategy === "naive",
    );
    return `| ${row.scenario} | ${row.paragraphs} | ${row.listItems} | ${row.strategy} | ${row.iterations} | ${row.medianAvgMs.toFixed(4)} | ${row.minAvgMs.toFixed(4)}-${row.maxAvgMs.toFixed(4)} | ${formatRelative(row, baseline)} | ${formatBytes(row.sizeBytes)} |`;
  }),
].join("\n");

const summary = `## DOCX Writer Benchmark

Measured in GitHub Actions with Bun on \`ubuntu-latest\`.

This benchmark compares a deliberately naive writer against an optimized writer. \`naive\` rebuilds XML strings as package entries, re-encodes each part, computes CRCs, creates many small ZIP buffers, and concatenates them. \`optimized\` caches stable OOXML part bytes on the Output Projection and writes the ZIP into one preallocated buffer.

Each row reports the median of ${SAMPLE_COUNT} samples. Each sample performs the listed number of writer iterations.

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

function runStrategy(
  strategy: DocxWriteStrategy,
  projection: ReturnType<typeof createPipelineFromLexicalJson>["outputProjection"],
  iterations: number,
  warmupIterations = 0,
): Sample {
  for (let index = 0; index < warmupIterations; index += 1) {
    createDocxBlob(projection, metadata, strategy);
  }

  const start = performance.now();
  let sizeBytes = 0;

  for (let index = 0; index < iterations; index += 1) {
    sizeBytes = createDocxBlob(projection, metadata, strategy).size;
  }

  return {
    durationMs: performance.now() - start,
    sizeBytes,
  };
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

function median(values: number[]): number {
  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);

  if (sorted.length % 2 === 1) {
    return sorted[middle] ?? 0;
  }

  return ((sorted[middle - 1] ?? 0) + (sorted[middle] ?? 0)) / 2;
}

function formatRelative(row: Row, baseline: Row | undefined): string {
  if (!baseline || row.strategy === "naive") {
    return "baseline";
  }

  const percent = ((row.medianAvgMs - baseline.medianAvgMs) / baseline.medianAvgMs) * 100;
  const direction = percent <= 0 ? "faster" : "slower";
  return `${Math.abs(percent).toFixed(1)}% ${direction}`;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  return `${(bytes / 1024).toFixed(1)} KB`;
}
