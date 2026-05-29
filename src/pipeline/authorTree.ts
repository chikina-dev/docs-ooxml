import type {
  AuthorBlock,
  AuthorInline,
  AuthorListItem,
  AuthorTree,
  HeadingLevel,
  ListKind,
  TextMarks,
} from "./types";

type LexicalTextFormat = number | string;
type LexicalHeadingTag = "h1" | "h2" | "h3";

export type LexicalNodeJson = {
  type: string;
  text?: string;
  format?: LexicalTextFormat;
  tag?: string;
  listType?: string;
  children?: LexicalNodeJson[];
};

export type LexicalEditorJson = {
  root: {
    children?: LexicalNodeJson[];
  };
};

type SerializedLexicalNode =
  | {
      type: "text";
      text: string;
      format?: LexicalTextFormat;
    }
  | {
      type: "linebreak";
    }
  | {
      type: "paragraph";
      children: SerializedLexicalNode[];
    }
  | {
      type: "heading";
      tag?: LexicalHeadingTag;
      children: SerializedLexicalNode[];
    }
  | {
      type: "list";
      listType: ListKind;
      children: SerializedLexicalNode[];
    }
  | {
      type: "listitem";
      children: SerializedLexicalNode[];
    }
  | {
      type: "unsupported";
      children: SerializedLexicalNode[];
    };

const TEXT_FORMAT = {
  bold: 1,
  italic: 1 << 1,
  underline: 1 << 3,
};

export function lexicalJsonToAuthorTree(json: LexicalEditorJson): AuthorTree {
  return {
    kind: "document",
    children: readEditorChildren(json).flatMap(readBlock),
  };
}

function readBlock(node: SerializedLexicalNode): AuthorBlock[] {
  if (node.type === "paragraph") {
    return [{ kind: "paragraph", children: readInlines(node.children ?? []) }];
  }

  if (node.type === "heading") {
    return [
      {
        kind: "heading",
        level: readHeadingLevel(node.tag),
        children: readInlines(node.children ?? []),
      },
    ];
  }

  if (node.type === "list") {
    const children = (node.children ?? []).flatMap((child): AuthorListItem[] => {
      if (child.type === "listitem") {
        return [{ kind: "listItem", children: readListItemInlines(child.children ?? []) }];
      }

      return [{ kind: "listItem", children: readInlines([child]) }];
    });
    return [{ kind: "list", listKind: node.listType, children }];
  }

  const children = readInlines([node]);
  return children.length > 0 ? [{ kind: "paragraph", children }] : [];
}

function readListItemInlines(children: SerializedLexicalNode[]): AuthorInline[] {
  return children.flatMap((child) => {
    if (child.type === "paragraph") {
      return readInlines(child.children ?? []);
    }

    if (child.type === "list") {
      return child.children.flatMap((nested) =>
        nested.type === "listitem" ? readListItemInlines(nested.children) : readInlines([nested]),
      );
    }

    return readInlines([child]);
  });
}

function readInlines(nodes: SerializedLexicalNode[]): AuthorInline[] {
  return nodes.flatMap((node): AuthorInline[] => {
    if (node.type === "text" && node.text !== undefined) {
      return [{ kind: "textRun", text: node.text, marks: readTextMarks(node.format) }];
    }

    if (node.type === "linebreak") {
      return [{ kind: "lineBreak" }];
    }

    return "children" in node ? readInlines(node.children) : [];
  });
}

function readHeadingLevel(tag: LexicalHeadingTag | undefined): HeadingLevel {
  if (tag === "h2") {
    return 2;
  }

  if (tag === "h3") {
    return 3;
  }

  return 1;
}

function readEditorChildren(json: LexicalEditorJson): SerializedLexicalNode[] {
  return readNodeList(json.root.children ?? []);
}

function readNodeList(value: LexicalNodeJson[]): SerializedLexicalNode[] {
  return value.map(readNode);
}

function readNode(value: LexicalNodeJson): SerializedLexicalNode {
  const originalType = value.type;
  if (originalType === "text" && typeof value.text === "string") {
    return { type: "text", text: value.text, format: readTextFormat(value.format) };
  }

  if (originalType === "linebreak") {
    return { type: "linebreak" };
  }

  if (originalType === "paragraph") {
    return { type: "paragraph", children: readNodeList(value.children ?? []) };
  }

  if (originalType === "heading") {
    return {
      type: "heading",
      tag: readHeadingTag(value.tag),
      children: readNodeList(value.children ?? []),
    };
  }

  if (originalType === "list") {
    return {
      type: "list",
      listType: readListKind(value.listType),
      children: readNodeList(value.children ?? []),
    };
  }

  if (originalType === "listitem") {
    return { type: "listitem", children: readNodeList(value.children ?? []) };
  }

  return { type: "unsupported", children: readNodeList(value.children ?? []) };
}

function readTextFormat(value: LexicalNodeJson["format"]): LexicalTextFormat | undefined {
  return typeof value === "number" || typeof value === "string" ? value : undefined;
}

function readHeadingTag(value: LexicalNodeJson["tag"]): LexicalHeadingTag | undefined {
  if (value === "h1" || value === "h2" || value === "h3") {
    return value;
  }

  return undefined;
}

function readListKind(value: LexicalNodeJson["listType"]): ListKind {
  return value === "number" ? "number" : "bullet";
}

function readTextMarks(format: LexicalTextFormat | undefined): TextMarks {
  if (typeof format === "string") {
    return {
      bold: format.includes("bold") || undefined,
      italic: format.includes("italic") || undefined,
      underline: format.includes("underline") || undefined,
    };
  }

  const bitmask = format ?? 0;
  return {
    bold: (bitmask & TEXT_FORMAT.bold) !== 0 || undefined,
    italic: (bitmask & TEXT_FORMAT.italic) !== 0 || undefined,
    underline: (bitmask & TEXT_FORMAT.underline) !== 0 || undefined,
  };
}
