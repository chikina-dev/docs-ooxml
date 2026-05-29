import type { EditorState, LexicalEditor, TextFormatType } from "lexical";
import {
  $createParagraphNode,
  $createTextNode,
  $getRoot,
  $getSelection,
  $isRangeSelection,
  CAN_REDO_COMMAND,
  CAN_UNDO_COMMAND,
  COMMAND_PRIORITY_LOW,
  FORMAT_TEXT_COMMAND,
  REDO_COMMAND,
  SELECTION_CHANGE_COMMAND,
  UNDO_COMMAND,
} from "lexical";
import { useCallback, useEffect, useMemo, useState } from "react";
import {
  ListItemNode,
  ListNode,
  INSERT_ORDERED_LIST_COMMAND,
  INSERT_UNORDERED_LIST_COMMAND,
  REMOVE_LIST_COMMAND,
} from "@lexical/list";
import { HeadingNode, $createHeadingNode } from "@lexical/rich-text";
import { $setBlocksType } from "@lexical/selection";
import { mergeRegister } from "@lexical/utils";
import { LexicalComposer } from "@lexical/react/LexicalComposer";
import { useLexicalComposerContext } from "@lexical/react/LexicalComposerContext";
import { RichTextPlugin } from "@lexical/react/LexicalRichTextPlugin";
import { ContentEditable } from "@lexical/react/LexicalContentEditable";
import { HistoryPlugin } from "@lexical/react/LexicalHistoryPlugin";
import { ListPlugin } from "@lexical/react/LexicalListPlugin";
import { OnChangePlugin } from "@lexical/react/LexicalOnChangePlugin";
import { LexicalErrorBoundary } from "@lexical/react/LexicalErrorBoundary";
import { createDocxBlob } from "./ooxml/docx.ts";
import { createPipelineFromLexicalJson } from "./pipeline/createPipeline.ts";
import type { OutputProjection } from "./pipeline/types.ts";

type PipelineState = ReturnType<typeof createPipelineFromLexicalJson>;
type HeadingLevel = 1 | 2 | 3;

const INITIAL_PIPELINE = createPipelineFromLexicalJson({
  root: {
    children: [
      {
        type: "paragraph",
        children: [{ type: "text", text: "Start writing, then export to Word.", format: 0 }],
      },
    ],
  },
});

const editorTheme = {
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

export default function App() {
  const [title, setTitle] = useState("Untitled Document");
  const [pipeline, setPipeline] = useState<PipelineState>(INITIAL_PIPELINE);

  const initialConfig = useMemo(
    () => ({
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
    }),
    [],
  );

  const handleChange = useCallback((editorState: EditorState) => {
    setPipeline(createPipelineFromLexicalJson(editorState.toJSON()));
  }, []);

  const handlePackageBufferExport = useCallback(() => {
    downloadDocx(pipeline.outputProjection, title, "package-buffer");
  }, [pipeline.outputProjection, title]);

  const handleChunkedZipExport = useCallback(() => {
    downloadDocx(pipeline.outputProjection, title, "chunked-zip");
  }, [pipeline.outputProjection, title]);

  return (
    <main className="app-shell">
      <section className="workspace">
        <header className="titlebar">
          <div className="brand-lockup">
            <span className="word-mark" aria-hidden="true">
              W
            </span>
            <div>
              <p className="eyebrow">Docs OOXML</p>
              <h1>Word authoring pipeline</h1>
            </div>
          </div>
          <label className="title-field">
            <span>Document title</span>
            <input value={title} onChange={(event) => setTitle(event.target.value)} />
          </label>
        </header>

        <LexicalComposer initialConfig={initialConfig}>
          <Toolbar
            onChunkedZipExport={handleChunkedZipExport}
            onPackageBufferExport={handlePackageBufferExport}
          />
          <div className="document-canvas">
            <div className="page-shadow">
              <div className="editor-frame">
                <RichTextPlugin
                  contentEditable={
                    <ContentEditable className="editor-input" aria-label="Document body" />
                  }
                  placeholder={<div className="editor-placeholder">Write the document body...</div>}
                  ErrorBoundary={LexicalErrorBoundary}
                />
                <HistoryPlugin />
                <ListPlugin />
                <OnChangePlugin onChange={handleChange} ignoreSelectionChange />
              </div>
            </div>
          </div>
        </LexicalComposer>
      </section>

      <PipelinePreview pipeline={pipeline} />
    </main>
  );
}

function Toolbar({
  onChunkedZipExport,
  onPackageBufferExport,
}: {
  onChunkedZipExport: () => void;
  onPackageBufferExport: () => void;
}) {
  const [editor] = useLexicalComposerContext();
  const [canUndo, setCanUndo] = useState(false);
  const [canRedo, setCanRedo] = useState(false);
  const [activeMarks, setActiveMarks] = useState({
    bold: false,
    italic: false,
    underline: false,
  });

  const refreshSelection = useCallback(() => {
    const selection = $getSelection();
    if ($isRangeSelection(selection)) {
      setActiveMarks({
        bold: selection.hasFormat("bold"),
        italic: selection.hasFormat("italic"),
        underline: selection.hasFormat("underline"),
      });
    }

    return false;
  }, []);

  useEffect(() => {
    return mergeRegister(
      editor.registerCommand(
        CAN_UNDO_COMMAND,
        (payload) => {
          setCanUndo(payload);
          return false;
        },
        COMMAND_PRIORITY_LOW,
      ),
      editor.registerCommand(
        CAN_REDO_COMMAND,
        (payload) => {
          setCanRedo(payload);
          return false;
        },
        COMMAND_PRIORITY_LOW,
      ),
      editor.registerCommand(SELECTION_CHANGE_COMMAND, refreshSelection, COMMAND_PRIORITY_LOW),
    );
  }, [editor, refreshSelection]);

  return (
    <div className="ribbon" aria-label="Editor toolbar">
      <nav className="ribbon-tabs" aria-label="Ribbon tabs">
        <button className="ribbon-tab is-active" type="button">
          Home
        </button>
        <button className="ribbon-tab" type="button">
          Insert
        </button>
        <button className="ribbon-tab" type="button">
          Layout
        </button>
        <button className="ribbon-tab" type="button">
          Review
        </button>
      </nav>
      <div className="ribbon-body">
        <div className="ribbon-group">
          <div className="ribbon-controls">
            <ToolbarButton
              label="Undo"
              disabled={!canUndo}
              onClick={() => editor.dispatchCommand(UNDO_COMMAND, undefined)}
            />
            <ToolbarButton
              label="Redo"
              disabled={!canRedo}
              onClick={() => editor.dispatchCommand(REDO_COMMAND, undefined)}
            />
          </div>
          <span className="ribbon-label">Clipboard</span>
        </div>
        <div className="ribbon-group">
          <div className="ribbon-controls">
            <ToolbarButton label="P" onClick={() => formatParagraph(editor)} />
            <ToolbarButton label="H1" onClick={() => formatHeading(editor, 1)} />
            <ToolbarButton label="H2" onClick={() => formatHeading(editor, 2)} />
            <ToolbarButton label="H3" onClick={() => formatHeading(editor, 3)} />
          </div>
          <span className="ribbon-label">Styles</span>
        </div>
        <div className="ribbon-group">
          <div className="ribbon-controls">
            <ToolbarButton
              label="B"
              active={activeMarks.bold}
              onClick={() => formatText(editor, "bold")}
            />
            <ToolbarButton
              label="I"
              active={activeMarks.italic}
              onClick={() => formatText(editor, "italic")}
            />
            <ToolbarButton
              label="U"
              active={activeMarks.underline}
              onClick={() => formatText(editor, "underline")}
            />
          </div>
          <span className="ribbon-label">Font</span>
        </div>
        <div className="ribbon-group">
          <div className="ribbon-controls">
            <ToolbarButton
              label="Bullets"
              onClick={() => editor.dispatchCommand(INSERT_UNORDERED_LIST_COMMAND, undefined)}
            />
            <ToolbarButton
              label="Numbers"
              onClick={() => editor.dispatchCommand(INSERT_ORDERED_LIST_COMMAND, undefined)}
            />
          </div>
          <span className="ribbon-label">Paragraph</span>
        </div>
        <div className="export-actions">
          <button className="export-button" type="button" onClick={onPackageBufferExport}>
            Export buffer
          </button>
          <button className="export-button secondary" type="button" onClick={onChunkedZipExport}>
            Export chunked
          </button>
        </div>
      </div>
    </div>
  );
}

function ToolbarButton({
  label,
  active = false,
  disabled = false,
  onClick,
}: {
  label: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      className={active ? "tool-button is-active" : "tool-button"}
      type="button"
      disabled={disabled}
      onClick={onClick}
    >
      {label}
    </button>
  );
}

function PipelinePreview({ pipeline }: { pipeline: PipelineState }) {
  return (
    <aside className="pipeline-panel" aria-label="Pipeline preview">
      <PreviewSection title="Author Tree" data={pipeline.authorTree} />
      <PreviewSection title="Semantic Author Graph" data={pipeline.semanticGraph} />
      <PreviewSection title="Output Projection" data={pipeline.outputProjection} />
    </aside>
  );
}

function PreviewSection({ title, data }: { title: string; data: unknown }) {
  return (
    <details className="preview-section" open>
      <summary>{title}</summary>
      <pre>{JSON.stringify(data, null, 2)}</pre>
    </details>
  );
}

function formatText(editor: LexicalEditor, format: TextFormatType) {
  editor.dispatchCommand(FORMAT_TEXT_COMMAND, format);
}

function formatParagraph(editor: LexicalEditor) {
  editor.dispatchCommand(REMOVE_LIST_COMMAND, undefined);
  editor.update(() => {
    const selection = $getSelection();
    if ($isRangeSelection(selection)) {
      $setBlocksType(selection, () => $createParagraphNode());
    }
  });
}

function formatHeading(editor: LexicalEditor, level: HeadingLevel) {
  editor.dispatchCommand(REMOVE_LIST_COMMAND, undefined);
  editor.update(() => {
    const selection = $getSelection();
    if ($isRangeSelection(selection)) {
      $setBlocksType(selection, () => $createHeadingNode(`h${level}` as "h1" | "h2" | "h3"));
    }
  });
}

function downloadDocx(
  projection: OutputProjection,
  title: string,
  strategy: "package-buffer" | "chunked-zip",
) {
  const safeTitle = slugifyFileName(title || "Untitled Document");
  const blob = createDocxBlob(
    projection,
    {
      title: title || "Untitled Document",
      creator: "Docs OOXML",
      createdAt: new Date(),
    },
    strategy,
  );
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeTitle}-${strategy}.docx`;
  link.click();
  URL.revokeObjectURL(url);
}

function slugifyFileName(value: string): string {
  return (
    value
      .trim()
      .replace(/[^\w.-]+/g, "-")
      .replace(/^-+|-+$/g, "") || "Untitled-Document"
  );
}
