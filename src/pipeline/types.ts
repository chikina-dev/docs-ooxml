export type TextMarks = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
};

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
      level: 1 | 2 | 3;
      children: AuthorInline[];
    }
  | {
      kind: "list";
      listKind: "bullet" | "number";
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
      id: string;
      sourcePath: string;
      sectionId: string;
      kind: "paragraph";
      inlines: SemanticInline[];
    }
  | {
      id: string;
      sourcePath: string;
      sectionId: string;
      createsSectionId: string;
      kind: "heading";
      level: 1 | 2 | 3;
      title: string;
      outlinePath: number[];
      inlines: SemanticInline[];
    }
  | {
      id: string;
      sourcePath: string;
      sectionId: string;
      listGroupId: string;
      ordinalInList: number;
      kind: "listItem";
      listKind: "bullet" | "number";
      level: number;
      inlines: SemanticInline[];
    };

export type SemanticSection = {
  id: string;
  parentId?: string;
  headingBlockId?: string;
  title: string;
  level: 0 | 1 | 2 | 3;
  outlinePath: number[];
  childSectionIds: string[];
  blockIds: string[];
};

export type SemanticListGroup = {
  id: string;
  sectionId: string;
  listKind: "bullet" | "number";
  itemIds: string[];
};

export type SemanticEdge = {
  from: string;
  to: string;
  kind: "contains" | "opensSection" | "continuesList";
};

export type SemanticAuthorGraph = {
  kind: "semanticAuthorGraph";
  rootSectionId: string;
  readingOrder: string[];
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
      sourceBlockId: string;
      sectionId: string;
      semanticRole: "body";
      styleId: "Normal";
      runs: WordRun[];
    }
  | {
      kind: "heading";
      sourceBlockId: string;
      sectionId: string;
      semanticRole: "sectionHeading";
      styleId: "Heading1" | "Heading2" | "Heading3";
      runs: WordRun[];
    }
  | {
      kind: "listParagraph";
      sourceBlockId: string;
      sectionId: string;
      semanticRole: "listItem";
      styleId: "ListParagraph";
      numberingRef: {
        listGroupId: string;
        listKind: "bullet" | "number";
        level: number;
        numId: 1 | 2;
      };
      level: number;
      runs: WordRun[];
    };

export type OutputProjection = {
  kind: "ooxmlPartProjection";
  target: "microsoftWordDocx";
  documentPlan: {
    page: "letter";
    styles: Array<"Normal" | "Heading1" | "Heading2" | "Heading3" | "ListParagraph">;
    numbering: Array<{
      listGroupId: string;
      listKind: "bullet" | "number";
      numId: 1 | 2;
    }>;
  };
  paragraphs: WordParagraph[];
  parts: OoxmlPartProjection[];
};

export type OoxmlPartProjection = {
  path: string;
  role:
    | "contentTypes"
    | "rootRelationships"
    | "coreProperties"
    | "appProperties"
    | "document"
    | "styles"
    | "numbering"
    | "documentRelationships";
  contentType?: string;
  xml: string;
};
