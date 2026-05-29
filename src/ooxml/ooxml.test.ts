import { describe, expect, it } from "vite-plus/test";
import { createPipelineFromLexicalJson } from "../pipeline/createPipeline";
import { benchmarkDocxWriters, createDocxBlob, createDocxPackage } from "./docx";
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

  it("emits Word blobs with both writer strategies", async () => {
    const naiveBlob = createDocxBlob(projection, metadata, "naive");
    const optimizedBlob = createDocxBlob(projection, metadata, "optimized");
    const bytes = new Uint8Array(await optimizedBlob.arrayBuffer());

    expect(naiveBlob.type).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    expect(naiveBlob.size).toBe(optimizedBlob.size);
    expect(listZipEntries(bytes)).toContain("word/document.xml");
    expect(listZipEntries(bytes)).toContain("word/numbering.xml");
  });

  it("benchmarks both docx writer strategies", () => {
    let time = 0;
    const results = benchmarkDocxWriters(projection, metadata, 2, () => {
      time += 3;
      return time;
    });

    expect(results).toEqual([
      { strategy: "naive", durationMs: 3, sizeBytes: expect.any(Number), iterations: 2 },
      { strategy: "optimized", durationMs: 3, sizeBytes: expect.any(Number), iterations: 2 },
    ]);
  });
});
