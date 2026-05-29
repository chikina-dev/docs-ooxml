import {
  createNaiveDocxBlob,
  createOptimizedDocxBlob,
  prepareDocxPackage,
  type DocxWriteStrategy,
} from "../src/ooxml/docx";
import { createPipelineFromLexicalJson } from "../src/pipeline/createPipeline";
import type { OutputProjection } from "../src/pipeline/types";

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

type BunRuntime = {
  env: Record<string, string | undefined>;
  write(path: string, content: string): Promise<number>;
};

declare global {
  var Bun: BunRuntime | undefined;
}

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
  const samples = new Map<DocxWriteStrategy, Sample[]>(
    STRATEGIES.map((strategy) => [strategy, []]),
  );

  for (const strategy of STRATEGIES) {
    runStrategy(strategy, scenario, 3, 0);
  }

  for (let sampleIndex = 0; sampleIndex < SAMPLE_COUNT; sampleIndex += 1) {
    const order = sampleIndex % 2 === 0 ? STRATEGIES : [...STRATEGIES].reverse();
    for (const strategy of order) {
      samples
        .get(strategy)
        ?.push(
          runStrategy(strategy, scenario, scenario.iterations, sampleIndex * scenario.iterations),
        );
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

This benchmark compares a deliberately naive writer against an optimized writer while generating a fresh Output Projection for every iteration. \`naive\` rebuilds XML strings as package entries, re-encodes each part, computes CRCs, creates many small ZIP buffers, and concatenates them. \`optimized\` prepares each fresh projection, reuses only projection-independent static OOXML parts, and writes the ZIP into one preallocated buffer.

Each row reports the median of ${SAMPLE_COUNT} samples. Each sample performs the listed number of writer iterations.

${table}
`;

console.log(summary);

const stepSummaryPath = globalThis.Bun?.env.GITHUB_STEP_SUMMARY;
if (stepSummaryPath) {
  await globalThis.Bun?.write(stepSummaryPath, `${summary}\n`);
}

function runStrategy(
  strategy: DocxWriteStrategy,
  scenario: Scenario,
  iterations: number,
  variantOffset: number,
): Sample {
  const start = performance.now();
  let sizeBytes = 0;

  for (let index = 0; index < iterations; index += 1) {
    const projection = createPipelineFromLexicalJson(
      createLexicalFixture(scenario, variantOffset + index),
    ).outputProjection;
    sizeBytes = createBlob(strategy, projection).size;
  }

  return {
    durationMs: performance.now() - start,
    sizeBytes,
  };
}

function createBlob(strategy: DocxWriteStrategy, projection: OutputProjection): Blob {
  if (strategy === "naive") {
    return createNaiveDocxBlob(projection, metadata);
  }

  const preparedPackage = prepareDocxPackage(projection);
  return createOptimizedDocxBlob(preparedPackage, metadata);
}

function createLexicalFixture(scenario: Scenario, variant = 0) {
  const variantLabel = variant.toString().padStart(6, "0");

  return {
    root: {
      children: [
        {
          type: "heading",
          tag: "h1",
          children: [
            { type: "text", text: `${scenario.name} benchmark ${variantLabel}`, format: 1 },
          ],
        },
        ...Array.from({ length: scenario.paragraphs }, (_, index) => ({
          type: "paragraph",
          children: [
            { type: "text", text: `Paragraph ${index + 1}: `, format: 1 },
            {
              type: "text",
              text: `This paragraph exercises XML escaping, run projection, and ZIP packaging for variant ${variantLabel}. `,
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
            children: [
              {
                type: "text",
                text: `Measured list item ${index + 1} for variant ${variantLabel}`,
                format: 0,
              },
            ],
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
