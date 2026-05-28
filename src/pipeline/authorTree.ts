import type { AuthorBlock, AuthorInline, AuthorListItem, AuthorTree, TextMarks } from "./types.ts";

type LexicalNodeJson = {
  type?: string;
  text?: string;
  format?: number | string;
  tag?: string;
  listType?: string;
  children?: LexicalNodeJson[];
};

type LexicalEditorJson = {
  root?: {
    children?: LexicalNodeJson[];
  };
};

const TEXT_FORMAT = {
  bold: 1,
  italic: 1 << 1,
  underline: 1 << 3,
};

export function lexicalJsonToAuthorTree(json: LexicalEditorJson): AuthorTree {
  return {
    kind: "document",
    children: (json.root?.children ?? []).flatMap(readBlock),
  };
}

function readBlock(node: LexicalNodeJson): AuthorBlock[] {
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
    const listKind = node.listType === "number" ? "number" : "bullet";
    const children = (node.children ?? []).flatMap((child): AuthorListItem[] => {
      if (child.type === "listitem") {
        return [{ kind: "listItem", children: readListItemInlines(child.children ?? []) }];
      }

      return [{ kind: "listItem", children: readInlines([child]) }];
    });
    return [{ kind: "list", listKind, children }];
  }

  const children = readInlines([node]);
  return children.length > 0 ? [{ kind: "paragraph", children }] : [];
}

function readListItemInlines(children: LexicalNodeJson[]): AuthorInline[] {
  return children.flatMap((child) => {
    if (child.type === "paragraph") {
      return readInlines(child.children ?? []);
    }

    if (child.type === "list") {
      return child.children?.flatMap((nested) => readListItemInlines(nested.children ?? [])) ?? [];
    }

    return readInlines([child]);
  });
}

function readInlines(nodes: LexicalNodeJson[]): AuthorInline[] {
  return nodes.flatMap((node): AuthorInline[] => {
    if (node.type === "text" && node.text !== undefined) {
      return [{ kind: "textRun", text: node.text, marks: readTextMarks(node.format) }];
    }

    if (node.type === "linebreak") {
      return [{ kind: "lineBreak" }];
    }

    return readInlines(node.children ?? []);
  });
}

function readHeadingLevel(tag: string | undefined): 1 | 2 | 3 {
  if (tag === "h2") {
    return 2;
  }

  if (tag === "h3") {
    return 3;
  }

  return 1;
}

function readTextMarks(format: number | string | undefined): TextMarks {
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
