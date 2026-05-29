import type {
  OoxmlPartProjection,
  OutputProjection,
  SemanticAuthorGraph,
  SemanticBlock,
  SemanticInline,
  WordParagraph,
  WordRun,
} from "./types.ts";
import {
  appPropertiesXml,
  contentTypesXml,
  corePropertiesTemplateXml,
  documentRelationshipsXml,
  documentXml,
  numberingXml,
  rootRelationshipsXml,
  stylesXml,
} from "../ooxml/parts.ts";

export function semanticGraphToOutputProjection(graph: SemanticAuthorGraph): OutputProjection {
  const numbering = graph.listGroups.map((listGroup) => ({
    listGroupId: listGroup.id,
    listKind: listGroup.listKind,
    numId: (listGroup.listKind === "number" ? 2 : 1) as 1 | 2,
  }));
  const paragraphs = graph.readingOrder.map((blockId) =>
    blockToWordParagraph(blockById(graph, blockId)),
  );
  const documentPlan = {
    page: "letter" as const,
    styles: ["Normal", "Heading1", "Heading2", "Heading3", "ListParagraph"] as Array<
      "Normal" | "Heading1" | "Heading2" | "Heading3" | "ListParagraph"
    >,
    numbering,
  };

  return {
    kind: "ooxmlPartProjection",
    target: "microsoftWordDocx",
    documentPlan,
    paragraphs,
    parts: createOoxmlPartProjection(paragraphs),
  };
}

function blockToWordParagraph(block: SemanticBlock): WordParagraph {
  if (block.kind === "heading") {
    return {
      kind: "heading",
      sourceBlockId: block.id,
      sectionId: block.createsSectionId,
      semanticRole: "sectionHeading",
      styleId: `Heading${block.level}` as const,
      runs: inlinesToRuns(block.inlines),
    };
  }

  if (block.kind === "listItem") {
    return {
      kind: "listParagraph",
      sourceBlockId: block.id,
      sectionId: block.sectionId,
      semanticRole: "listItem",
      styleId: "ListParagraph",
      numberingRef: {
        listGroupId: block.listGroupId,
        listKind: block.listKind,
        level: block.level,
        numId: block.listKind === "number" ? 2 : 1,
      },
      level: block.level,
      runs: inlinesToRuns(block.inlines),
    };
  }

  return {
    kind: "paragraph",
    sourceBlockId: block.id,
    sectionId: block.sectionId,
    semanticRole: "body",
    styleId: "Normal",
    runs: inlinesToRuns(block.inlines),
  };
}

function createOoxmlPartProjection(paragraphs: WordParagraph[]): OoxmlPartProjection[] {
  return [
    {
      path: "[Content_Types].xml",
      role: "contentTypes",
      xml: contentTypesXml(),
    },
    {
      path: "_rels/.rels",
      role: "rootRelationships",
      xml: rootRelationshipsXml(),
    },
    {
      path: "docProps/core.xml",
      role: "coreProperties",
      contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      xml: corePropertiesTemplateXml(),
    },
    {
      path: "docProps/app.xml",
      role: "appProperties",
      contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      xml: appPropertiesXml(),
    },
    {
      path: "word/document.xml",
      role: "document",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      xml: documentXml(paragraphs),
    },
    {
      path: "word/styles.xml",
      role: "styles",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      xml: stylesXml(),
    },
    {
      path: "word/numbering.xml",
      role: "numbering",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
      xml: numberingXml(),
    },
    {
      path: "word/_rels/document.xml.rels",
      role: "documentRelationships",
      xml: documentRelationshipsXml(),
    },
  ];
}

function blockById(graph: SemanticAuthorGraph, blockId: string): SemanticBlock {
  const block = graph.blocks.find((item) => item.id === blockId);
  if (!block) {
    throw new Error(`Output projection could not find semantic block ${blockId}`);
  }

  return block;
}

function inlinesToRuns(inlines: SemanticInline[]): WordRun[] {
  return inlines.map((inline) => {
    if (inline.kind === "break") {
      return { kind: "break" };
    }

    return {
      kind: "text",
      text: inline.text,
      marks: { ...inline.marks },
    };
  });
}
