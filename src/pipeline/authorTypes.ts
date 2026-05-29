export type TextMarks = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
};

export type HeadingLevel = 1 | 2 | 3;
export type ListKind = "bullet" | "number";
export type AuthorSourcePath = `/children/${number}` | `/children/${number}/children/${number}`;

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
