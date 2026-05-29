import type { OoxmlPartProjection, OutputProjection } from "../pipeline/types.ts";
import { corePropertiesXml } from "./parts.ts";
import { createZip, createZipBlobParts, type ZipEntry } from "./zip.ts";

export type DocxMetadata = {
  title: string;
  creator: string;
  createdAt: Date;
};

export type DocxWriteStrategy = "package-buffer" | "chunked-zip";

export type DocxBenchmarkResult = {
  strategy: DocxWriteStrategy;
  durationMs: number;
  sizeBytes: number;
  iterations: number;
};

const WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export function createDocxBlob(
  projection: OutputProjection,
  metadata: DocxMetadata,
  strategy: DocxWriteStrategy = "package-buffer",
): Blob {
  if (strategy === "chunked-zip") {
    return createDocxBlobChunkedZip(projection, metadata);
  }

  return createDocxBlobPackageBuffer(projection, metadata);
}

export function createDocxBlobPackageBuffer(
  projection: OutputProjection,
  metadata: DocxMetadata,
): Blob {
  const zip = createZip(createDocxPackage(projection, metadata));
  const buffer = new ArrayBuffer(zip.byteLength);
  new Uint8Array(buffer).set(zip);

  return new Blob([buffer], { type: WORD_MIME });
}

export function createDocxBlobChunkedZip(
  projection: OutputProjection,
  metadata: DocxMetadata,
): Blob {
  return new Blob(createZipBlobParts(createDocxPackage(projection, metadata)), { type: WORD_MIME });
}

export function benchmarkDocxWriters(
  projection: OutputProjection,
  metadata: DocxMetadata,
  iterations = 50,
  now: () => number = () => performance.now(),
): DocxBenchmarkResult[] {
  return (["package-buffer", "chunked-zip"] as const).map((strategy) => {
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

function partXml(part: OoxmlPartProjection, metadata: DocxMetadata): string {
  if (part.role === "coreProperties") {
    return corePropertiesXml(metadata);
  }

  return part.xml;
}
