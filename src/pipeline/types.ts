export type TextMarks = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
};

export type HeadingLevel = 1 | 2 | 3;
export type SectionLevel = 0 | HeadingLevel;
export type ListKind = "bullet" | "number";

export type SemanticBlockId = `block-${number}`;
export type SemanticSectionId = "section-root" | `section-${number}`;
export type SemanticListGroupId = `list-${number}`;
export type SemanticNodeId = SemanticBlockId | SemanticSectionId | SemanticListGroupId;
export type AuthorSourcePath = `/children/${number}` | `/children/${number}/children/${number}`;

export type WordBodyStyleId = "Normal";
export type WordHeadingStyleId = "Heading1" | "Heading2" | "Heading3";
export type WordListStyleId = "ListParagraph";
export type WordStyleId = WordBodyStyleId | WordHeadingStyleId | WordListStyleId;
export type WordNumberingId = 1 | 2;

export type AuthorInline =
  | {
      kind: "textRun";
      text: string;
      marks: TextMarks;
    }
  | {
      kind: "lineBreak";
    };

export type AuthorBlock =
  | {
      kind: "paragraph";
      children: AuthorInline[];
    }
  | {
      kind: "heading";
      level: HeadingLevel;
      children: AuthorInline[];
    }
  | {
      kind: "list";
      listKind: ListKind;
      children: AuthorListItem[];
    };

export type AuthorListItem = {
  kind: "listItem";
  children: AuthorInline[];
};

export type AuthorTree = {
  kind: "document";
  children: AuthorBlock[];
};

export type SemanticInline =
  | {
      kind: "span";
      text: string;
      marks: TextMarks;
    }
  | {
      kind: "break";
    };

export type SemanticBlock =
  | {
      id: SemanticBlockId;
      sourcePath: AuthorSourcePath;
      sectionId: SemanticSectionId;
      kind: "paragraph";
      inlines: SemanticInline[];
    }
  | {
      id: SemanticBlockId;
      sourcePath: AuthorSourcePath;
      sectionId: SemanticSectionId;
      createsSectionId: SemanticSectionId;
      kind: "heading";
      level: HeadingLevel;
      title: string;
      outlinePath: number[];
      inlines: SemanticInline[];
    }
  | {
      id: SemanticBlockId;
      sourcePath: AuthorSourcePath;
      sectionId: SemanticSectionId;
      listGroupId: SemanticListGroupId;
      ordinalInList: number;
      kind: "listItem";
      listKind: ListKind;
      level: number;
      inlines: SemanticInline[];
    };

export type SemanticSection = {
  id: SemanticSectionId;
  parentId?: SemanticSectionId;
  headingBlockId?: SemanticBlockId;
  title: string;
  level: SectionLevel;
  outlinePath: number[];
  childSectionIds: SemanticSectionId[];
  blockIds: SemanticBlockId[];
};

export type SemanticListGroup = {
  id: SemanticListGroupId;
  sectionId: SemanticSectionId;
  listKind: ListKind;
  itemIds: SemanticBlockId[];
};

export type SemanticEdge = {
  from: SemanticNodeId;
  to: SemanticNodeId;
  kind: "contains" | "opensSection" | "continuesList";
};

export type SemanticAuthorGraph = {
  kind: "semanticAuthorGraph";
  rootSectionId: SemanticSectionId;
  readingOrder: SemanticBlockId[];
  sections: SemanticSection[];
  listGroups: SemanticListGroup[];
  blocks: SemanticBlock[];
  edges: SemanticEdge[];
};

export type WordRun =
  | {
      kind: "text";
      text: string;
      marks: TextMarks;
    }
  | {
      kind: "break";
    };

export type WordParagraph =
  | {
      kind: "paragraph";
      sourceBlockId: SemanticBlockId;
      sectionId: SemanticSectionId;
      semanticRole: "body";
      styleId: WordBodyStyleId;
      runs: WordRun[];
    }
  | {
      kind: "heading";
      sourceBlockId: SemanticBlockId;
      sectionId: SemanticSectionId;
      semanticRole: "sectionHeading";
      styleId: WordHeadingStyleId;
      runs: WordRun[];
    }
  | {
      kind: "listParagraph";
      sourceBlockId: SemanticBlockId;
      sectionId: SemanticSectionId;
      semanticRole: "listItem";
      styleId: WordListStyleId;
      numberingRef: {
        listGroupId: SemanticListGroupId;
        listKind: ListKind;
        level: number;
        numId: WordNumberingId;
      };
      level: number;
      runs: WordRun[];
    };

export type OutputProjection = {
  kind: "ooxmlPartProjection";
  target: "microsoftWordDocx";
  documentPlan: {
    page: "letter";
    styles: readonly WordStyleId[];
    numbering: readonly {
      listGroupId: SemanticListGroupId;
      listKind: ListKind;
      numId: WordNumberingId;
    }[];
  };
  paragraphs: WordParagraph[];
  parts: OoxmlPartProjectionList;
};

export type OoxmlPackagePath =
  | "[Content_Types].xml"
  | "_rels/.rels"
  | "docProps/core.xml"
  | "docProps/app.xml"
  | "word/document.xml"
  | "word/styles.xml"
  | "word/numbering.xml"
  | "word/_rels/document.xml.rels";

export type ContentTypesPartProjection = {
  path: "[Content_Types].xml";
  role: "contentTypes";
  contentType?: never;
  xml: string;
};

export type RootRelationshipsPartProjection = {
  path: "_rels/.rels";
  role: "rootRelationships";
  contentType?: never;
  xml: string;
};

export type CorePropertiesPartProjection = {
  path: "docProps/core.xml";
  role: "coreProperties";
  contentType: "application/vnd.openxmlformats-package.core-properties+xml";
  xml: string;
};

export type AppPropertiesPartProjection = {
  path: "docProps/app.xml";
  role: "appProperties";
  contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml";
  xml: string;
};

export type DocumentPartProjection = {
  path: "word/document.xml";
  role: "document";
  contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
  xml: string;
};

export type StylesPartProjection = {
  path: "word/styles.xml";
  role: "styles";
  contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
  xml: string;
};

export type NumberingPartProjection = {
  path: "word/numbering.xml";
  role: "numbering";
  contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
  xml: string;
};

export type DocumentRelationshipsPartProjection = {
  path: "word/_rels/document.xml.rels";
  role: "documentRelationships";
  contentType?: never;
  xml: string;
};

export type OoxmlPartProjection =
  | ContentTypesPartProjection
  | RootRelationshipsPartProjection
  | CorePropertiesPartProjection
  | AppPropertiesPartProjection
  | DocumentPartProjection
  | StylesPartProjection
  | NumberingPartProjection
  | DocumentRelationshipsPartProjection;

export type StaticOoxmlPartProjection = Exclude<
  OoxmlPartProjection,
  CorePropertiesPartProjection | DocumentPartProjection
>;

export type OoxmlPartProjectionList = readonly [
  ContentTypesPartProjection,
  RootRelationshipsPartProjection,
  CorePropertiesPartProjection,
  AppPropertiesPartProjection,
  DocumentPartProjection,
  StylesPartProjection,
  NumberingPartProjection,
  DocumentRelationshipsPartProjection,
];
