import type { OoxmlPartProjection, OutputProjection } from "../pipeline/types.ts";
import { corePropertiesXml } from "./parts.ts";
import {
  createStoredZipEntry,
  createZipFromStoredEntries,
  createZipNaive,
  type StoredZipEntry,
  type ZipEntry,
} from "./zip.ts";

export type DocxMetadata = {
  title: string;
  creator: string;
  createdAt: Date;
};

export type DocxWriteStrategy = "naive" | "optimized";

export type PreparedDocxPackage = {
  kind: "preparedDocxPackage";
  stableEntries: StoredZipEntry[];
};

export type DocxBenchmarkResult = {
  strategy: DocxWriteStrategy;
  durationMs: number;
  sizeBytes: number;
  iterations: number;
};

const WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const staticEntryCache = new Map<string, StoredZipEntry>();

export function createDocxBlob(
  projection: OutputProjection,
  metadata: DocxMetadata,
  strategy: DocxWriteStrategy = "optimized",
): Blob {
  if (strategy === "naive") {
    return createNaiveDocxBlob(projection, metadata);
  }

  return createOptimizedDocxBlob(prepareDocxPackage(projection), metadata);
}

export function prepareDocxPackage(projection: OutputProjection): PreparedDocxPackage {
  return {
    kind: "preparedDocxPackage",
    stableEntries: projection.parts
      .filter((part) => part.role !== "coreProperties")
      .map((part) =>
        part.role === "document"
          ? createStoredZipEntry(part.path, part.xml)
          : storedStaticEntry(part),
      ),
  };
}

export function createOptimizedDocxBlob(
  preparedPackage: PreparedDocxPackage,
  metadata: DocxMetadata,
): Blob {
  const zip = createZipFromStoredEntries([
    ...preparedPackage.stableEntries,
    createStoredZipEntry("docProps/core.xml", corePropertiesXml(metadata)),
  ]);

  return new Blob([zip.buffer as ArrayBuffer], { type: WORD_MIME });
}

export function createNaiveDocxBlob(projection: OutputProjection, metadata: DocxMetadata): Blob {
  const zip = createZipNaive(createDocxPackage(projection, metadata));

  return new Blob([zip.buffer as ArrayBuffer], { type: WORD_MIME });
}

export function benchmarkDocxWriters(
  projection: OutputProjection,
  metadata: DocxMetadata,
  iterations = 50,
  now: () => number = () => performance.now(),
): DocxBenchmarkResult[] {
  const preparedPackage = prepareDocxPackage(projection);

  return (["naive", "optimized"] as const).map((strategy) => {
    const start = now();
    let sizeBytes = 0;

    for (let index = 0; index < iterations; index += 1) {
      sizeBytes =
        strategy === "naive"
          ? createNaiveDocxBlob(projection, metadata).size
          : createOptimizedDocxBlob(preparedPackage, metadata).size;
    }

    return {
      strategy,
      durationMs: now() - start,
      sizeBytes,
      iterations,
    };
  });
}

export function createDocxPackage(
  projection: OutputProjection,
  metadata: DocxMetadata,
): ZipEntry[] {
  return projection.parts.map((part) => ({
    path: part.path,
    data: partXml(part, metadata),
  }));
}

function partXml(part: OoxmlPartProjection, metadata: DocxMetadata): string {
  if (part.role === "coreProperties") {
    return corePropertiesXml(metadata);
  }

  return part.xml;
}

function storedStaticEntry(part: OoxmlPartProjection): StoredZipEntry {
  const cached = staticEntryCache.get(part.path);
  if (cached) {
    return cached;
  }

  const entry = createStoredZipEntry(part.path, part.xml);
  staticEntryCache.set(part.path, entry);
  return entry;
}
