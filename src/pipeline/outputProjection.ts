import type {
  OutputProjection,
  SemanticAuthorGraph,
  SemanticBlock,
  SemanticInline,
  WordRun,
} from "./types.ts";

export function semanticGraphToOutputProjection(graph: SemanticAuthorGraph): OutputProjection {
  const numbering = graph.listGroups.map((listGroup) => ({
    listGroupId: listGroup.id,
    listKind: listGroup.listKind,
    numId: (listGroup.listKind === "number" ? 2 : 1) as 1 | 2,
  }));

  return {
    kind: "wordOutputProjection",
    target: "microsoftWord",
    documentPlan: {
      page: "letter",
      styles: ["Normal", "Heading1", "Heading2", "Heading3", "ListParagraph"],
      numbering,
    },
    paragraphs: graph.readingOrder.map((blockId) =>
      blockToWordParagraph(blockById(graph, blockId)),
    ),
  };
}

function blockToWordParagraph(block: SemanticBlock): OutputProjection["paragraphs"][number] {
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
