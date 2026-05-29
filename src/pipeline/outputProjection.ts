import type {
  HeadingLevel,
  ListKind,
  OoxmlPartProjectionList,
  OutputProjection,
  SemanticAuthorGraph,
  SemanticBlockId,
  SemanticBlock,
  SemanticInline,
  WordHeadingStyleId,
  WordNumberingId,
  WordParagraph,
  WordRun,
  WordStyleId,
} from "./types";
import {
  appPropertiesXml,
  contentTypesXml,
  corePropertiesTemplateXml,
  documentRelationshipsXml,
  documentXml,
  numberingXml,
  rootRelationshipsXml,
  stylesXml,
} from "../ooxml/parts";

const WORD_STYLES: readonly WordStyleId[] = [
  "Normal",
  "Heading1",
  "Heading2",
  "Heading3",
  "ListParagraph",
];

export function semanticGraphToOutputProjection(graph: SemanticAuthorGraph): OutputProjection {
  const numbering = graph.listGroups.map((listGroup) => ({
    listGroupId: listGroup.id,
    listKind: listGroup.listKind,
    numId: numberingIdForListKind(listGroup.listKind),
  }));
  const paragraphs = graph.readingOrder.map((blockId) =>
    blockToWordParagraph(blockById(graph, blockId)),
  );
  const documentPlan: OutputProjection["documentPlan"] = {
    page: "letter",
    styles: WORD_STYLES,
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
      styleId: headingStyleId(block.level),
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
        numId: numberingIdForListKind(block.listKind),
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

function createOoxmlPartProjection(paragraphs: WordParagraph[]): OoxmlPartProjectionList {
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
  ] satisfies OoxmlPartProjectionList;
}

function blockById(graph: SemanticAuthorGraph, blockId: SemanticBlockId): SemanticBlock {
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

function headingStyleId(level: HeadingLevel): WordHeadingStyleId {
  if (level === 1) {
    return "Heading1";
  }

  if (level === 2) {
    return "Heading2";
  }

  return "Heading3";
}

function numberingIdForListKind(listKind: ListKind): WordNumberingId {
  return listKind === "number" ? 2 : 1;
}
