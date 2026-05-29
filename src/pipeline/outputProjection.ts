import type {
  HeadingLevel,
  ListKind,
  OoxmlPartProjectionList,
  OoxmlTextTagOpen,
  OutputProjection,
  SemanticAuthorGraph,
  SemanticBlockId,
  SemanticBlock,
  SemanticInline,
  TextMarks,
  WordHeadingStyleId,
  WordNumberingId,
  WordParagraph,
  WordRun,
  WordStyleId,
} from "./types";
import { escapeXml } from "./xmlEscaping";

const WORD_STYLES: readonly WordStyleId[] = [
  "Normal",
  "Heading1",
  "Heading2",
  "Heading3",
  "ListParagraph",
];
const HEADING_PARAGRAPH_PROPERTIES: Record<WordHeadingStyleId, string> = {
  Heading1: '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>',
  Heading2: '<w:pPr><w:pStyle w:val="Heading2"/></w:pPr>',
  Heading3: '<w:pPr><w:pStyle w:val="Heading3"/></w:pPr>',
};
const RUN_PROPERTIES_BY_MARK_MASK = [
  "",
  "<w:rPr><w:b/></w:rPr>",
  "<w:rPr><w:i/></w:rPr>",
  "<w:rPr><w:b/><w:i/></w:rPr>",
  '<w:rPr><w:u w:val="single"/></w:rPr>',
  '<w:rPr><w:b/><w:u w:val="single"/></w:rPr>',
  '<w:rPr><w:i/><w:u w:val="single"/></w:rPr>',
  '<w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr>',
] as const;

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
    const styleId = headingStyleId(block.level);
    return {
      kind: "heading",
      sourceBlockId: block.id,
      sectionId: block.createsSectionId,
      semanticRole: "sectionHeading",
      styleId,
      ooxml: {
        paragraphPropertiesXml: HEADING_PARAGRAPH_PROPERTIES[styleId],
      },
      runs: inlinesToRuns(block.inlines),
    };
  }

  if (block.kind === "listItem") {
    const numId = numberingIdForListKind(block.listKind);
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
        numId,
      },
      level: block.level,
      ooxml: {
        paragraphPropertiesXml: `<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${block.level}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`,
      },
      runs: inlinesToRuns(block.inlines),
    };
  }

  return {
    kind: "paragraph",
    sourceBlockId: block.id,
    sectionId: block.sectionId,
    semanticRole: "body",
    styleId: "Normal",
    ooxml: {
      paragraphPropertiesXml: "",
    },
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
      return {
        kind: "break",
        ooxml: {
          runXml: "<w:r><w:br/></w:r>",
        },
      };
    }

    return {
      kind: "text",
      text: inline.text,
      marks: { ...inline.marks },
      ooxml: {
        runPropertiesXml: RUN_PROPERTIES_BY_MARK_MASK[textMarkMask(inline.marks)] ?? "",
        textTagOpen: textTagOpen(inline.text),
        escapedText: escapeXml(inline.text),
      },
    };
  });
}

function textMarkMask(marks: TextMarks): number {
  let mask = 0;

  if (marks.bold) {
    mask |= 1;
  }

  if (marks.italic) {
    mask |= 2;
  }

  if (marks.underline) {
    mask |= 4;
  }

  return mask;
}

function textTagOpen(value: string): OoxmlTextTagOpen {
  return /^\s|\s$/.test(value) ? '<w:t xml:space="preserve">' : "<w:t>";
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
