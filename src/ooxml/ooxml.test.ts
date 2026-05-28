import { describe, expect, it } from "vite-plus/test";
import type { OutputProjection } from "../pipeline/types.ts";
import { createDocxBlob, createDocxPackage } from "./docx.ts";
import { escapeXml } from "./xml.ts";
import { crc32, createZip, listZipEntries } from "./zip.ts";

const projection: OutputProjection = {
  kind: "wordOutputProjection",
  target: "microsoftWord",
  documentPlan: {
    page: "letter",
    styles: ["Normal", "Heading1", "Heading2", "Heading3", "ListParagraph"],
    numbering: [{ listGroupId: "list-1", listKind: "bullet", numId: 1 }],
  },
  paragraphs: [
    {
      kind: "heading",
      sourceBlockId: "block-1",
      sectionId: "section-1",
      semanticRole: "sectionHeading",
      styleId: "Heading1",
      runs: [{ kind: "text", text: "Title & Intro", marks: { bold: true } }],
    },
    {
      kind: "paragraph",
      sourceBlockId: "block-2",
      sectionId: "section-1",
      semanticRole: "body",
      styleId: "Normal",
      runs: [
        { kind: "text", text: "Hello <world>", marks: { italic: true, underline: true } },
        { kind: "break" },
        { kind: "text", text: "Done", marks: {} },
      ],
    },
    {
      kind: "listParagraph",
      sourceBlockId: "block-3",
      sectionId: "section-1",
      semanticRole: "listItem",
      styleId: "ListParagraph",
      numberingRef: {
        listGroupId: "list-1",
        listKind: "bullet",
        level: 0,
        numId: 1,
      },
      level: 0,
      runs: [{ kind: "text", text: "Point", marks: {} }],
    },
  ],
};

const metadata = {
  title: "A&B <Doc>",
  creator: "Docs OOXML",
  createdAt: new Date("2026-05-28T00:00:00.000Z"),
};

describe("OOXML writer", () => {
  it("escapes XML text and metadata", () => {
    expect(escapeXml(`A&B <"doc">'`)).toBe("A&amp;B &lt;&quot;doc&quot;&gt;&apos;");

    const documentXml = createDocxPackage(projection, metadata).find(
      (entry) => entry.path === "word/document.xml",
    )?.data;
    expect(documentXml).toContain("Title &amp; Intro");
    expect(documentXml).toContain("Hello &lt;world&gt;");
  });

  it("creates the required docx package entries", () => {
    const packageEntries = createDocxPackage(projection, metadata);

    expect(packageEntries.map((entry) => entry.path)).toEqual([
      "[Content_Types].xml",
      "_rels/.rels",
      "docProps/core.xml",
      "docProps/app.xml",
      "word/document.xml",
      "word/styles.xml",
      "word/numbering.xml",
      "word/_rels/document.xml.rels",
    ]);
  });

  it("writes an uncompressed zip with valid headers, directory, and CRC32", () => {
    const zip = createZip([
      { path: "a.txt", data: "hello" },
      { path: "b.txt", data: "world" },
    ]);

    expect(zip[0]).toBe(0x50);
    expect(listZipEntries(zip)).toEqual(["a.txt", "b.txt"]);
    expect(crc32(new TextEncoder().encode("hello"))).toBe(0x3610a686);
    expect(zip.at(-22)).toBe(0x50);
  });

  it("emits a Word blob with required OOXML entries", async () => {
    const blob = createDocxBlob(projection, metadata);
    const bytes = new Uint8Array(await blob.arrayBuffer());

    expect(blob.type).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    expect(listZipEntries(bytes)).toContain("word/document.xml");
    expect(listZipEntries(bytes)).toContain("word/numbering.xml");
  });
});
