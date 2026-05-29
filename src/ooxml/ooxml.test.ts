import { describe, expect, it } from "vite-plus/test";
import { createPipelineFromLexicalJson } from "../pipeline/createPipeline";
import { unzipSync } from "fflate";
import { benchmarkDocxWriters, createDocxBlob, createDocxPackage } from "./docx";
import { documentXml, documentXmlOptimized } from "./parts";
import { escapeXml } from "./xml";
import { crc32, createZipNaive, createZipOptimized, listZipEntries } from "./zip";

const projection = createPipelineFromLexicalJson({
  root: {
    children: [
      {
        type: "heading",
        tag: "h1",
        children: [{ type: "text", text: "Title & Intro", format: 1 }],
      },
      {
        type: "paragraph",
        children: [
          { type: "text", text: "Hello <world>", format: 10 },
          { type: "linebreak" },
          { type: "text", text: "Done", format: 0 },
        ],
      },
      {
        type: "list",
        listType: "bullet",
        children: [{ type: "listitem", children: [{ type: "text", text: "Point", format: 0 }] }],
      },
    ],
  },
}).outputProjection;

const metadata = {
  title: "A&B <Doc>",
  creator: "Docs OOXML",
  createdAt: new Date("2026-05-28T00:00:00.000Z"),
};

describe("OOXML writer", () => {
  it("projects output as OOXML package parts", () => {
    expect(projection).toMatchObject({
      kind: "ooxmlPartProjection",
      target: "microsoftWordDocx",
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
  });

  it("projects OOXML-ready paragraph and run fragments", () => {
    expect(projection.paragraphs[0]).toMatchObject({
      kind: "heading",
      ooxml: {
        paragraphPropertiesXml: '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>',
      },
      runs: [
        {
          kind: "text",
          ooxml: {
            runPropertiesXml: "<w:rPr><w:b/></w:rPr>",
            textTagOpen: "<w:t>",
            escapedText: "Title &amp; Intro",
          },
        },
      ],
    });
    expect(projection.paragraphs[1]?.runs[0]).toMatchObject({
      kind: "text",
      ooxml: {
        runPropertiesXml: '<w:rPr><w:i/><w:u w:val="single"/></w:rPr>',
        escapedText: "Hello &lt;world&gt;",
      },
    });
    expect(projection.paragraphs[2]).toMatchObject({
      kind: "listParagraph",
      ooxml: {
        paragraphPropertiesXml:
          '<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>',
      },
    });
  });

  it("escapes XML text and metadata", () => {
    expect(escapeXml(`A&B <"doc">'`)).toBe("A&amp;B &lt;&quot;doc&quot;&gt;&apos;");

    const documentXml = createDocxPackage(projection, metadata).find(
      (entry) => entry.path === "word/document.xml",
    )?.data;
    const coreXml = createDocxPackage(projection, metadata).find(
      (entry) => entry.path === "docProps/core.xml",
    )?.data;
    expect(documentXml).toContain("Title &amp; Intro");
    expect(documentXml).toContain("Hello &lt;world&gt;");
    expect(coreXml).toContain("A&amp;B &lt;Doc&gt;");
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

  it("keeps optimized document XML equivalent to the naive document XML", () => {
    const documentPart = projection.parts.find((part) => part.role === "document");

    expect(documentPart?.role).toBe("document");
    if (documentPart?.role !== "document") {
      return;
    }

    expect(documentXmlOptimized(documentPart.paragraphs)).toBe(
      documentXml(documentPart.paragraphs),
    );
  });

  it("writes equivalent uncompressed zips with naive and optimized writers", () => {
    const entries = [
      { path: "a.txt", data: "hello" },
      { path: "b.txt", data: "world" },
    ];
    const naiveZip = createZipNaive(entries);
    const optimizedZip = createZipOptimized(entries);

    expect(naiveZip).toEqual(optimizedZip);
    expect(listZipEntries(optimizedZip)).toEqual(["a.txt", "b.txt"]);
    expect(crc32(new TextEncoder().encode("hello"))).toBe(0x3610a686);
    expect(optimizedZip.at(-22)).toBe(0x50);
  });

  it("emits Word blobs with all writer strategies", async () => {
    const naiveBlob = createDocxBlob(projection, metadata, "naive");
    const optimizedBlob = createDocxBlob(projection, metadata, "optimized");
    const fflateBlob = createDocxBlob(projection, metadata, "fflate-store");
    const fflateStreamBlob = createDocxBlob(projection, metadata, "fflate-stream");
    const optimizedBytes = new Uint8Array(await optimizedBlob.arrayBuffer());
    const fflateBytes = new Uint8Array(await fflateBlob.arrayBuffer());
    const fflateStreamBytes = new Uint8Array(await fflateStreamBlob.arrayBuffer());

    expect(naiveBlob.type).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    expect(naiveBlob.size).toBe(optimizedBlob.size);
    expect(fflateBlob.type).toBe(naiveBlob.type);
    expect(fflateStreamBlob.type).toBe(naiveBlob.type);
    expect(listZipEntries(optimizedBytes)).toContain("word/document.xml");
    expect(listZipEntries(optimizedBytes)).toContain("word/numbering.xml");
    expect(listZipEntries(fflateBytes)).toContain("word/document.xml");
    expect(listZipEntries(fflateBytes)).toContain("word/numbering.xml");
    expect(Object.keys(unzipSync(fflateStreamBytes))).toContain("word/document.xml");
    expect(Object.keys(unzipSync(fflateStreamBytes))).toContain("word/numbering.xml");
  });

  it("benchmarks all docx writer strategies", () => {
    let time = 0;
    const results = benchmarkDocxWriters(projection, metadata, 2, () => {
      time += 3;
      return time;
    });

    expect(results).toEqual([
      { strategy: "naive", durationMs: 3, sizeBytes: expect.any(Number), iterations: 2 },
      { strategy: "optimized", durationMs: 3, sizeBytes: expect.any(Number), iterations: 2 },
      { strategy: "fflate-store", durationMs: 3, sizeBytes: expect.any(Number), iterations: 2 },
      { strategy: "fflate-stream", durationMs: 3, sizeBytes: expect.any(Number), iterations: 2 },
    ]);
  });
});
