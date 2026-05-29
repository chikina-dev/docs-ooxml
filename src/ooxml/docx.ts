import type {
  CorePropertiesPartProjection,
  DocumentPartProjection,
  OoxmlPackagePath,
  OoxmlPartProjection,
  OutputProjection,
  StaticOoxmlPartProjection,
} from "../pipeline/outputProjectionTypes";
import { strToU8, Zip, ZipPassThrough, zipSync, type Zippable } from "fflate";
import {
  appPropertiesXml,
  contentTypesXml,
  corePropertiesXml,
  documentRelationshipsXml,
  documentXml,
  documentXmlOptimized,
  numberingXml,
  rootRelationshipsXml,
  stylesXml,
  writeDocumentXmlChunks,
} from "./parts";
import {
  createStoredZipEntry,
  createZipFromStoredEntries,
  createZipNaive,
  type StoredZipEntry,
  type ZipEntry,
} from "./zip";

export type DocxMetadata = {
  title: string;
  creator: string;
  createdAt: Date;
};

export type DocxWriteStrategy = "naive" | "optimized" | "fflate-store" | "fflate-stream";

export type PreparedDocxPackage = {
  kind: "preparedDocxPackage";
  stableEntries: readonly StoredZipEntry<OoxmlPackagePath>[];
};

export type DocxWriter = {
  strategy: DocxWriteStrategy;
  createBlob: (projection: OutputProjection, metadata: DocxMetadata) => Blob;
  createBenchmarkBlob?: (
    projection: OutputProjection,
    metadata: DocxMetadata,
    preparedPackage: PreparedDocxPackage,
  ) => Blob;
};

export type DocxBenchmarkResult = {
  strategy: DocxWriteStrategy;
  durationMs: number;
  sizeBytes: number;
  iterations: number;
};

type StaticPartRole = StaticOoxmlPartProjection["role"];

const WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
export const DEFAULT_DOCX_WRITE_STRATEGY: DocxWriteStrategy = "fflate-store";
export const DOCX_WRITERS: readonly DocxWriter[] = [
  {
    strategy: "naive",
    createBlob: createNaiveDocxBlob,
  },
  {
    strategy: "optimized",
    createBlob: (projection, metadata) =>
      createOptimizedDocxBlob(prepareDocxPackage(projection), metadata),
    createBenchmarkBlob: (_projection, metadata, preparedPackage) =>
      createOptimizedDocxBlob(preparedPackage, metadata),
  },
  {
    strategy: "fflate-store",
    createBlob: createFflateStoreDocxBlob,
  },
  {
    strategy: "fflate-stream",
    createBlob: createFflateStreamDocxBlob,
  },
];
const FFLATE_MTIME = new Date("1980-01-01T00:00:00.000Z");
const FFLATE_STORE_OPTIONS: { readonly level: 0; readonly mtime: Date } = {
  level: 0,
  mtime: FFLATE_MTIME,
};
const DOCUMENT_XML_STREAM_CHARS = 64 * 1024;
const STATIC_ENTRY_BY_ROLE: Record<StaticPartRole, StoredZipEntry<OoxmlPackagePath>> = {
  contentTypes: createStoredZipEntry("[Content_Types].xml", contentTypesXml()),
  rootRelationships: createStoredZipEntry("_rels/.rels", rootRelationshipsXml()),
  appProperties: createStoredZipEntry("docProps/app.xml", appPropertiesXml()),
  styles: createStoredZipEntry("word/styles.xml", stylesXml()),
  numbering: createStoredZipEntry("word/numbering.xml", numberingXml()),
  documentRelationships: createStoredZipEntry(
    "word/_rels/document.xml.rels",
    documentRelationshipsXml(),
  ),
};

export function createDocxBlob(
  projection: OutputProjection,
  metadata: DocxMetadata,
  strategy: DocxWriteStrategy = DEFAULT_DOCX_WRITE_STRATEGY,
): Blob {
  return docxWriterForStrategy(strategy).createBlob(projection, metadata);
}

export function prepareDocxPackage(projection: OutputProjection): PreparedDocxPackage {
  return {
    kind: "preparedDocxPackage",
    stableEntries: projection.parts
      .filter((part) => !isCorePropertiesPart(part))
      .map((part) =>
        isDocumentPart(part)
          ? createStoredZipEntry(part.path, documentXmlOptimized(part.paragraphs))
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

  return new Blob([zip], { type: WORD_MIME });
}

export function createNaiveDocxBlob(projection: OutputProjection, metadata: DocxMetadata): Blob {
  const zip = createZipNaive(createDocxPackage(projection, metadata));

  return new Blob([zip], { type: WORD_MIME });
}

export function createFflateStoreDocxBlob(
  projection: OutputProjection,
  metadata: DocxMetadata,
): Blob {
  const zip = zipSync(createFflateStorePackage(projection, metadata), FFLATE_STORE_OPTIONS);

  return new Blob([zip], { type: WORD_MIME });
}

export function createFflateStreamDocxBlob(
  projection: OutputProjection,
  metadata: DocxMetadata,
): Blob {
  const zip = createFflateStreamStoreZip(projection, metadata);

  return new Blob([zip], { type: WORD_MIME });
}

export function benchmarkDocxWriters(
  projection: OutputProjection,
  metadata: DocxMetadata,
  iterations = 50,
  now: () => number = () => performance.now(),
): DocxBenchmarkResult[] {
  const preparedPackage = prepareDocxPackage(projection);

  return DOCX_WRITERS.map((writer) => {
    const start = now();
    let sizeBytes = 0;

    for (let index = 0; index < iterations; index += 1) {
      sizeBytes = (writer.createBenchmarkBlob ?? writer.createBlob)(
        projection,
        metadata,
        preparedPackage,
      ).size;
    }

    return {
      strategy: writer.strategy,
      durationMs: now() - start,
      sizeBytes,
      iterations,
    };
  });
}

export function docxWriterForStrategy(strategy: DocxWriteStrategy): DocxWriter {
  const writer = DOCX_WRITERS.find((candidate) => candidate.strategy === strategy);
  if (!writer) {
    throw new Error(`Unknown DOCX write strategy: ${strategy}`);
  }

  return writer;
}

export function createDocxPackage(
  projection: OutputProjection,
  metadata: DocxMetadata,
): ZipEntry<OoxmlPackagePath>[] {
  return projection.parts.map((part) => ({
    path: part.path,
    data: partXml(part, metadata),
  }));
}

function partXml(part: OoxmlPartProjection, metadata: DocxMetadata): string {
  if (part.role === "coreProperties") {
    return corePropertiesXml(metadata);
  }

  if (part.role === "document") {
    return documentXml(part.paragraphs);
  }

  return staticPartXml(part);
}

function isCorePropertiesPart(part: OoxmlPartProjection): part is CorePropertiesPartProjection {
  return part.role === "coreProperties";
}

function isDocumentPart(part: OoxmlPartProjection): part is DocumentPartProjection {
  return part.role === "document";
}

function storedStaticEntry(part: StaticOoxmlPartProjection): StoredZipEntry<OoxmlPackagePath> {
  return STATIC_ENTRY_BY_ROLE[part.role];
}

function createFflateStorePackage(projection: OutputProjection, metadata: DocxMetadata): Zippable {
  const entries: Zippable = {};

  for (const part of projection.parts) {
    entries[part.path] = [partDataBytes(part, metadata), FFLATE_STORE_OPTIONS];
  }

  return entries;
}

function createFflateStreamStoreZip(
  projection: OutputProjection,
  metadata: DocxMetadata,
): Uint8Array<ArrayBuffer> {
  const chunks: Uint8Array[] = [];
  const zip = new Zip((error, chunk) => {
    if (error) {
      throw error;
    }

    chunks.push(chunk);
  });

  for (const part of projection.parts) {
    const entry = new ZipPassThrough(part.path);
    entry.mtime = FFLATE_MTIME;
    zip.add(entry);

    if (part.role === "document") {
      pushDocumentXmlToEntry(entry, part.paragraphs);
    } else {
      entry.push(partDataBytes(part, metadata), true);
    }
  }

  zip.end();
  return concatBytes(chunks);
}

function pushDocumentXmlToEntry(
  entry: ZipPassThrough,
  paragraphs: DocumentPartProjection["paragraphs"],
): void {
  const pendingChunks: string[] = [];
  let pendingLength = 0;

  writeDocumentXmlChunks(paragraphs, (chunk) => {
    pendingChunks.push(chunk);
    pendingLength += chunk.length;

    if (pendingLength >= DOCUMENT_XML_STREAM_CHARS) {
      flushDocumentXmlChunks(entry, pendingChunks, false);
      pendingLength = 0;
    }
  });

  flushDocumentXmlChunks(entry, pendingChunks, true);
}

function flushDocumentXmlChunks(
  entry: ZipPassThrough,
  pendingChunks: string[],
  final: boolean,
): void {
  if (pendingChunks.length === 0) {
    if (final) {
      entry.push(new Uint8Array(), true);
    }
    return;
  }

  entry.push(strToU8(pendingChunks.join("")), final);
  pendingChunks.length = 0;
}

function partDataBytes(part: OoxmlPartProjection, metadata: DocxMetadata): Uint8Array {
  if (part.role === "coreProperties") {
    return strToU8(corePropertiesXml(metadata));
  }

  if (part.role === "document") {
    return strToU8(documentXmlOptimized(part.paragraphs));
  }

  return storedStaticEntry(part).dataBytes;
}

function concatBytes(chunks: readonly Uint8Array[]): Uint8Array<ArrayBuffer> {
  const output = new Uint8Array(chunks.reduce((size, chunk) => size + chunk.length, 0));
  let offset = 0;

  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.length;
  }

  return output;
}

function staticPartXml(part: StaticOoxmlPartProjection): string {
  switch (part.role) {
    case "contentTypes":
      return contentTypesXml();
    case "rootRelationships":
      return rootRelationshipsXml();
    case "appProperties":
      return appPropertiesXml();
    case "styles":
      return stylesXml();
    case "numbering":
      return numberingXml();
    case "documentRelationships":
      return documentRelationshipsXml();
  }
}
