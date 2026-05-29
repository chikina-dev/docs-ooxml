import type { WordParagraph, WordRun } from "../pipeline/types";
import { escapeXml, xmlDeclaration } from "./xml";

export type CorePropertiesValues = {
  title: string;
  creator: string;
  createdAt: Date;
};

export type XmlChunkWriter = (chunk: string) => void;

export function contentTypesXml(): string {
  return `${xmlDeclaration()}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;
}

export function rootRelationshipsXml(): string {
  return `${xmlDeclaration()}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

export function documentRelationshipsXml(): string {
  return `${xmlDeclaration()}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;
}

export function corePropertiesXml(values: CorePropertiesValues): string {
  const createdAt = values.createdAt.toISOString();
  return `${xmlDeclaration()}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(values.title)}</dc:title>
  <dc:creator>${escapeXml(values.creator)}</dc:creator>
  <cp:lastModifiedBy>${escapeXml(values.creator)}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${createdAt}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${createdAt}</dcterms:modified>
</cp:coreProperties>`;
}

export function corePropertiesTemplateXml(): string {
  return corePropertiesXml({
    title: "{{title}}",
    creator: "{{creator}}",
    createdAt: new Date("1970-01-01T00:00:00.000Z"),
  });
}

export function appPropertiesXml(): string {
  return `${xmlDeclaration()}
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Docs OOXML</Application>
</Properties>`;
}

export function documentXml(paragraphs: readonly WordParagraph[]): string {
  const body = paragraphs.map(paragraphXml).join("");
  return `${xmlDeclaration()}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`;
}

export function documentXmlOptimized(paragraphs: readonly WordParagraph[]): string {
  const chunks: string[] = [];
  writeDocumentXmlChunks(paragraphs, (chunk) => chunks.push(chunk));

  return chunks.join("");
}

export function writeDocumentXmlChunks(
  paragraphs: readonly WordParagraph[],
  write: XmlChunkWriter,
): void {
  write(`${xmlDeclaration()}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    `);

  for (const paragraph of paragraphs) {
    writeParagraphXml(write, paragraph);
  }

  write(`
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`);
}

export function stylesXml(): string {
  return `${xmlDeclaration()}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="480" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
</w:styles>`;
}

export function numberingXml(): string {
  return `${xmlDeclaration()}
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
  <w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`;
}

function paragraphXml(paragraph: WordParagraph): string {
  return `<w:p>${paragraph.ooxml.paragraphPropertiesXml}${paragraph.runs.map(runXml).join("")}</w:p>`;
}

function runXml(run: WordRun): string {
  if (run.kind === "break") {
    return run.ooxml.runXml;
  }

  return `<w:r>${run.ooxml.runPropertiesXml}${run.ooxml.textTagOpen}${run.ooxml.escapedText}</w:t></w:r>`;
}

function writeParagraphXml(write: XmlChunkWriter, paragraph: WordParagraph): void {
  write("<w:p>");
  write(paragraph.ooxml.paragraphPropertiesXml);

  for (const run of paragraph.runs) {
    writeRunXml(write, run);
  }

  write("</w:p>");
}

function writeRunXml(write: XmlChunkWriter, run: WordRun): void {
  if (run.kind === "break") {
    write(run.ooxml.runXml);
    return;
  }

  write("<w:r>");
  write(run.ooxml.runPropertiesXml);
  write(run.ooxml.textTagOpen);
  write(run.ooxml.escapedText);
  write("</w:t></w:r>");
}
