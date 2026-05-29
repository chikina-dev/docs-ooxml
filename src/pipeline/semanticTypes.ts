import type { AuthorSourcePath, HeadingLevel, ListKind, TextMarks } from "./authorTypes";

export type SectionLevel = 0 | HeadingLevel;
export type SemanticBlockId = `block-${number}`;
export type SemanticSectionId = "section-root" | `section-${number}`;
export type SemanticListGroupId = `list-${number}`;
export type SemanticNodeId = SemanticBlockId | SemanticSectionId | SemanticListGroupId;

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
