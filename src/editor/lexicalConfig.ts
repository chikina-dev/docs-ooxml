import { $createParagraphNode, $createTextNode, $getRoot, type EditorThemeClasses } from "lexical";
import { ListItemNode, ListNode } from "@lexical/list";
import { HeadingNode } from "@lexical/rich-text";
import type { InitialConfigType } from "@lexical/react/LexicalComposer";

const editorTheme: EditorThemeClasses = {
  paragraph: "editor-paragraph",
  heading: {
    h1: "editor-heading editor-heading-h1",
    h2: "editor-heading editor-heading-h2",
    h3: "editor-heading editor-heading-h3",
  },
  list: {
    nested: {
      listitem: "editor-nested-listitem",
    },
    ol: "editor-list editor-list-ol",
    ul: "editor-list editor-list-ul",
    listitem: "editor-listitem",
  },
  text: {
    bold: "editor-text-bold",
    italic: "editor-text-italic",
    underline: "editor-text-underline",
  },
};

export const lexicalInitialConfig: InitialConfigType = {
  namespace: "docs-ooxml-authoring",
  nodes: [HeadingNode, ListNode, ListItemNode],
  editorState() {
    const paragraph = $createParagraphNode();
    paragraph.append($createTextNode("Start writing, then export to Word."));
    $getRoot().append(paragraph);
  },
  onError(error: Error) {
    throw error;
  },
  theme: editorTheme,
};
