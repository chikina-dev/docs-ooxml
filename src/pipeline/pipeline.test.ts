import { describe, expect, it } from "vite-plus/test";
import { lexicalJsonToAuthorTree } from "./authorTree.ts";
import { semanticGraphToOutputProjection } from "./outputProjection.ts";
import { authorTreeToSemanticGraph } from "./semanticGraph.ts";

const lexicalJson = {
  root: {
    children: [
      {
        type: "heading",
        tag: "h1",
        children: [{ type: "text", text: "Product Spec", format: 1 }],
      },
      {
        type: "heading",
        tag: "h2",
        children: [{ type: "text", text: "Overview", format: 0 }],
      },
      {
        type: "paragraph",
        children: [
          { type: "text", text: "A ", format: 0 },
          { type: "text", text: "marked", format: 10 },
          { type: "linebreak" },
          { type: "unknown", children: [{ type: "text", text: " tail", format: 0 }] },
        ],
      },
      {
        type: "list",
        listType: "number",
        children: [
          {
            type: "listitem",
            children: [{ type: "text", text: "First", format: 0 }],
          },
          {
            type: "listitem",
            children: [{ type: "text", text: "Second", format: 0 }],
          },
        ],
      },
    ],
  },
};

describe("document pipeline", () => {
  it("converts Lexical JSON into an Author Tree without Word-specific concerns", () => {
    const tree = lexicalJsonToAuthorTree(lexicalJson);

    expect(tree.children[0]).toMatchObject({
      kind: "heading",
      level: 1,
      children: [{ kind: "textRun", text: "Product Spec", marks: { bold: true } }],
    });
    expect(tree.children[2]).toMatchObject({
      kind: "paragraph",
      children: [
        { kind: "textRun", text: "A " },
        { kind: "textRun", text: "marked", marks: { italic: true, underline: true } },
        { kind: "lineBreak" },
        { kind: "textRun", text: " tail" },
      ],
    });
    expect(tree.children[3]).toMatchObject({
      kind: "list",
      listKind: "number",
      children: [
        { kind: "listItem", children: [{ kind: "textRun", text: "First" }] },
        { kind: "listItem", children: [{ kind: "textRun", text: "Second" }] },
      ],
    });
  });

  it("turns author blocks into a section/list semantic graph", () => {
    const graph = authorTreeToSemanticGraph(lexicalJsonToAuthorTree(lexicalJson));

    expect(graph.rootSectionId).toBe("section-root");
    expect(graph.readingOrder).toEqual(["block-1", "block-2", "block-3", "block-4", "block-5"]);
    expect(graph.sections).toMatchObject([
      { id: "section-root", level: 0, childSectionIds: ["section-1"] },
      {
        id: "section-1",
        parentId: "section-root",
        headingBlockId: "block-1",
        title: "Product Spec",
        level: 1,
        outlinePath: [1],
        childSectionIds: ["section-2"],
      },
      {
        id: "section-2",
        parentId: "section-1",
        headingBlockId: "block-2",
        title: "Overview",
        level: 2,
        outlinePath: [1, 1],
        blockIds: ["block-3", "block-4", "block-5"],
      },
    ]);
    expect(graph.blocks).toMatchObject([
      {
        id: "block-1",
        kind: "heading",
        sectionId: "section-root",
        createsSectionId: "section-1",
        outlinePath: [1],
      },
      {
        id: "block-2",
        kind: "heading",
        sectionId: "section-1",
        createsSectionId: "section-2",
        outlinePath: [1, 1],
      },
      { id: "block-3", kind: "paragraph", sectionId: "section-2" },
      {
        id: "block-4",
        kind: "listItem",
        sectionId: "section-2",
        listGroupId: "list-1",
        ordinalInList: 1,
      },
      {
        id: "block-5",
        kind: "listItem",
        sectionId: "section-2",
        listGroupId: "list-1",
        ordinalInList: 2,
      },
    ]);
    expect(graph.listGroups).toEqual([
      { id: "list-1", sectionId: "section-2", listKind: "number", itemIds: ["block-4", "block-5"] },
    ]);
    expect(graph.edges).toContainEqual({ from: "block-2", to: "section-2", kind: "opensSection" });
    expect(graph.edges).toContainEqual({ from: "list-1", to: "block-5", kind: "continuesList" });
  });

  it("projects the semantic graph into OOXML part instructions", () => {
    const projection = semanticGraphToOutputProjection(
      authorTreeToSemanticGraph(lexicalJsonToAuthorTree(lexicalJson)),
    );

    expect(projection).toMatchObject({
      kind: "ooxmlPartProjection",
      target: "microsoftWordDocx",
      documentPlan: {
        page: "letter",
        styles: ["Normal", "Heading1", "Heading2", "Heading3", "ListParagraph"],
        numbering: [{ listGroupId: "list-1", listKind: "number", numId: 2 }],
      },
      parts: [
        { path: "[Content_Types].xml", role: "contentTypes" },
        { path: "_rels/.rels", role: "rootRelationships" },
        { path: "docProps/core.xml", role: "coreProperties" },
        { path: "docProps/app.xml", role: "appProperties" },
        { path: "word/document.xml", role: "document" },
        { path: "word/styles.xml", role: "styles" },
        { path: "word/numbering.xml", role: "numbering" },
        { path: "word/_rels/document.xml.rels", role: "documentRelationships" },
      ],
    });
    expect(projection.paragraphs).toMatchObject([
      {
        kind: "heading",
        sourceBlockId: "block-1",
        sectionId: "section-1",
        semanticRole: "sectionHeading",
        styleId: "Heading1",
      },
      {
        kind: "heading",
        sourceBlockId: "block-2",
        sectionId: "section-2",
        semanticRole: "sectionHeading",
        styleId: "Heading2",
      },
      {
        kind: "paragraph",
        sourceBlockId: "block-3",
        sectionId: "section-2",
        semanticRole: "body",
        styleId: "Normal",
      },
      {
        kind: "listParagraph",
        sourceBlockId: "block-4",
        sectionId: "section-2",
        semanticRole: "listItem",
        styleId: "ListParagraph",
        numberingRef: { listGroupId: "list-1", listKind: "number", level: 0, numId: 2 },
      },
      {
        kind: "listParagraph",
        sourceBlockId: "block-5",
        sectionId: "section-2",
        semanticRole: "listItem",
        styleId: "ListParagraph",
        numberingRef: { listGroupId: "list-1", listKind: "number", level: 0, numId: 2 },
      },
    ]);
  });
});
