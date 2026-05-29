import type {
  CorePropertiesPartProjection,
  DocumentPartProjection,
  OoxmlPackagePath,
  OoxmlPartProjection,
  OutputProjection,
  StaticOoxmlPartProjection,
} from "../pipeline/types";
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

export type DocxWriteStrategy = "naive" | "optimized";

export type PreparedDocxPackage = {
  kind: "preparedDocxPackage";
  stableEntries: readonly StoredZipEntry<OoxmlPackagePath>[];
};

export type DocxBenchmarkResult = {
  strategy: DocxWriteStrategy;
  durationMs: number;
  sizeBytes: number;
  iterations: number;
};

type StaticPartRole = StaticOoxmlPartProjection["role"];

const WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DOCX_WRITE_STRATEGIES: readonly DocxWriteStrategy[] = ["naive", "optimized"];
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

export function benchmarkDocxWriters(
  projection: OutputProjection,
  metadata: DocxMetadata,
  iterations = 50,
  now: () => number = () => performance.now(),
): DocxBenchmarkResult[] {
  const preparedPackage = prepareDocxPackage(projection);

  return DOCX_WRITE_STRATEGIES.map((strategy) => {
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
