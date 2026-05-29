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

export type DocxBenchmarkResult = {
  strategy: DocxWriteStrategy;
  durationMs: number;
  sizeBytes: number;
  iterations: number;
};

const WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const projectionEntryCache = new WeakMap<OutputProjection, StoredZipEntry[]>();

export function createDocxBlob(
  projection: OutputProjection,
  metadata: DocxMetadata,
  strategy: DocxWriteStrategy = "optimized",
): Blob {
  const zip =
    strategy === "naive"
      ? createZipNaive(createDocxPackage(projection, metadata))
      : createZipFromStoredEntries(createOptimizedDocxPackage(projection, metadata));

  return new Blob([zip.buffer as ArrayBuffer], { type: WORD_MIME });
}

export function benchmarkDocxWriters(
  projection: OutputProjection,
  metadata: DocxMetadata,
  iterations = 50,
  now: () => number = () => performance.now(),
): DocxBenchmarkResult[] {
  return (["naive", "optimized"] as const).map((strategy) => {
    const start = now();
    let sizeBytes = 0;

    for (let index = 0; index < iterations; index += 1) {
      sizeBytes = createDocxBlob(projection, metadata, strategy).size;
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

function createOptimizedDocxPackage(
  projection: OutputProjection,
  metadata: DocxMetadata,
): StoredZipEntry[] {
  return [
    ...cachedProjectionEntries(projection),
    createStoredZipEntry("docProps/core.xml", corePropertiesXml(metadata)),
  ];
}

function cachedProjectionEntries(projection: OutputProjection): StoredZipEntry[] {
  const cached = projectionEntryCache.get(projection);
  if (cached) {
    return cached;
  }

  const entries = projection.parts
    .filter((part) => part.role !== "coreProperties")
    .map((part) => createStoredZipEntry(part.path, part.xml));
  projectionEntryCache.set(projection, entries);

  return entries;
}

function partXml(part: OoxmlPartProjection, metadata: DocxMetadata): string {
  if (part.role === "coreProperties") {
    return corePropertiesXml(metadata);
  }

  return part.xml;
}
