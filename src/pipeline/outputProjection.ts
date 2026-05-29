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
  const blocksById = new Map(graph.blocks.map((block) => [block.id, block]));
  const paragraphs = graph.readingOrder.map((blockId) =>
    blockToWordParagraph(blockById(blocksById, blockId)),
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
      source: "staticXml",
    },
    {
      path: "_rels/.rels",
      role: "rootRelationships",
      source: "staticXml",
    },
    {
      path: "docProps/core.xml",
      role: "coreProperties",
      contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      source: "metadata",
    },
    {
      path: "docProps/app.xml",
      role: "appProperties",
      contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      source: "staticXml",
    },
    {
      path: "word/document.xml",
      role: "document",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      source: "wordParagraphs",
      paragraphs,
    },
    {
      path: "word/styles.xml",
      role: "styles",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      source: "staticXml",
    },
    {
      path: "word/numbering.xml",
      role: "numbering",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
      source: "staticXml",
    },
    {
      path: "word/_rels/document.xml.rels",
      role: "documentRelationships",
      source: "staticXml",
    },
  ] satisfies OoxmlPartProjectionList;
}

function blockById(
  blocksById: ReadonlyMap<SemanticBlockId, SemanticBlock>,
  blockId: SemanticBlockId,
): SemanticBlock {
  const block = blocksById.get(blockId);
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
