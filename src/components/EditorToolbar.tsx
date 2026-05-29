import type { LexicalEditor, TextFormatType } from "lexical";
import {
  $createParagraphNode,
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
import { useCallback, useEffect, useState } from "react";
import {
  INSERT_ORDERED_LIST_COMMAND,
  INSERT_UNORDERED_LIST_COMMAND,
  REMOVE_LIST_COMMAND,
} from "@lexical/list";
import { $createHeadingNode, type HeadingTagType } from "@lexical/rich-text";
import { $setBlocksType } from "@lexical/selection";
import { useLexicalComposerContext } from "@lexical/react/LexicalComposerContext";
import { mergeRegister } from "@lexical/utils";
import type { DocxWriteStrategy } from "../ooxml/docx";

type HeadingLevel = 1 | 2 | 3;

export function EditorToolbar({ onExport }: { onExport: (strategy: DocxWriteStrategy) => void }) {
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
          <button
            className="export-button secondary"
            type="button"
            onClick={() => onExport("naive")}
          >
            Export naive
          </button>
          <button className="export-button" type="button" onClick={() => onExport("optimized")}>
            Export optimized
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
      $setBlocksType(selection, () => $createHeadingNode(headingTagForLevel(level)));
    }
  });
}

function headingTagForLevel(level: HeadingLevel): HeadingTagType {
  if (level === 1) {
    return "h1";
  }

  if (level === 2) {
    return "h2";
  }

  return "h3";
}
