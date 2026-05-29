import { createDocxBlob, DOCX_WRITERS, type DocxWriteStrategy } from "../src/ooxml/docx";
import { createPipelineFromLexicalJson } from "../src/pipeline/createPipeline";
import type { OutputProjection } from "../src/pipeline/outputProjectionTypes";

type Scenario = {
  name: string;
  paragraphs: number;
  listItems: number;
  iterations: number;
};

type BenchmarkMode = {
  name: "single-run" | "batch";
  sampleCount: number;
  iterationsForScenario: (scenario: Scenario) => number;
};

type Sample = {
  durationMs: number;
  sizeBytes: number;
};

type Row = {
  mode: BenchmarkMode["name"];
  scenario: string;
  paragraphs: number;
  listItems: number;
  strategy: DocxWriteStrategy;
  samples: number;
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

const BENCHMARK_MODES: BenchmarkMode[] = [
  { name: "single-run", sampleCount: 21, iterationsForScenario: () => 1 },
  { name: "batch", sampleCount: 7, iterationsForScenario: (scenario) => scenario.iterations },
];
const STRATEGIES = DOCX_WRITERS.map((writer) => writer.strategy);

const metadata = {
  title: "GitHub Actions DOCX Writer Benchmark",
  creator: "Docs OOXML CI",
  createdAt: new Date("2026-05-29T00:00:00.000Z"),
};

const rows = BENCHMARK_MODES.flatMap((mode) =>
  SCENARIOS.flatMap((scenario) => benchmarkScenario(mode, scenario)),
);

function benchmarkScenario(mode: BenchmarkMode, scenario: Scenario): Row[] {
  const iterations = mode.iterationsForScenario(scenario);
  const samples = new Map<DocxWriteStrategy, Sample[]>(
    STRATEGIES.map((strategy) => [strategy, []]),
  );

  for (const strategy of STRATEGIES) {
    runStrategy(strategy, scenario, iterations, 0);
  }

  for (let sampleIndex = 0; sampleIndex < mode.sampleCount; sampleIndex += 1) {
    const order = sampleIndex % 2 === 0 ? STRATEGIES : [...STRATEGIES].reverse();
    for (const strategy of order) {
      samples
        .get(strategy)
        ?.push(runStrategy(strategy, scenario, iterations, sampleIndex * iterations));
    }
  }

  return STRATEGIES.map((strategy): Row => {
    const strategySamples = samples.get(strategy) ?? [];
    const averages = strategySamples.map((sample) => sample.durationMs / iterations);
    return {
      mode: mode.name,
      scenario: scenario.name,
      paragraphs: scenario.paragraphs,
      listItems: scenario.listItems,
      strategy,
      samples: mode.sampleCount,
      iterations,
      medianAvgMs: median(averages),
      minAvgMs: Math.min(...averages),
      maxAvgMs: Math.max(...averages),
      sizeBytes: strategySamples.at(-1)?.sizeBytes ?? 0,
    };
  });
}

const table = [
  "| Mode | Scenario | Paragraphs | List Items | Strategy | Samples | Iterations/sample | Median avg ms/run | Min-Max avg | vs naive | DOCX size |",
  "| --- | --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |",
  ...rows.map((row) => {
    const baseline = rows.find(
      (item) =>
        item.mode === row.mode && item.scenario === row.scenario && item.strategy === "naive",
    );
    return `| ${row.mode} | ${row.scenario} | ${row.paragraphs} | ${row.listItems} | ${row.strategy} | ${row.samples} | ${row.iterations} | ${row.medianAvgMs.toFixed(4)} | ${row.minAvgMs.toFixed(4)}-${row.maxAvgMs.toFixed(4)} | ${formatRelative(row, baseline)} | ${formatBytes(row.sizeBytes)} |`;
  }),
].join("\n");

const summary = `## DOCX Writer Benchmark

Measured in GitHub Actions with Bun on \`ubuntu-latest\`.

This benchmark compares a deliberately naive writer, the manual optimized writer, and fflate-backed store writers. Each measured iteration packages a different pre-generated Output Projection, so the writer never packages the same projection twice. \`single-run\` measures one package operation per sample; \`batch\` measures repeated packaging in one sample. \`naive\` materializes every OOXML part, re-encodes each part, computes CRCs, creates many small ZIP buffers, and concatenates them. \`optimized\` materializes only dynamic parts per projection, uses module-eager static OOXML entries from the first writer call, and writes the ZIP into one preallocated buffer. \`fflate-store\` uses fflate \`zipSync\` with compression disabled. \`fflate-stream\` writes \`document.xml\` in XML chunks into fflate \`ZipPassThrough\` with compression disabled.

Each row reports the median of its listed sample count. Each sample performs the listed number of writer iterations.

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
  const projections = Array.from(
    { length: iterations },
    (_, index) =>
      createPipelineFromLexicalJson(createLexicalFixture(scenario, variantOffset + index))
        .outputProjection,
  );
  const start = performance.now();
  let sizeBytes = 0;

  for (let index = 0; index < iterations; index += 1) {
    const projection = projections[index];
    if (projection) {
      sizeBytes = createBlob(strategy, projection).size;
    }
  }

  return {
    durationMs: performance.now() - start,
    sizeBytes,
  };
}

function createBlob(strategy: DocxWriteStrategy, projection: OutputProjection): Blob {
  return createDocxBlob(projection, metadata, strategy);
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
