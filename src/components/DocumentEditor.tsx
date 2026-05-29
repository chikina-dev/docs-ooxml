import type { EditorState } from "lexical";
import { ContentEditable } from "@lexical/react/LexicalContentEditable";
import { LexicalErrorBoundary } from "@lexical/react/LexicalErrorBoundary";
import { HistoryPlugin } from "@lexical/react/LexicalHistoryPlugin";
import { ListPlugin } from "@lexical/react/LexicalListPlugin";
import { OnChangePlugin } from "@lexical/react/LexicalOnChangePlugin";
import { RichTextPlugin } from "@lexical/react/LexicalRichTextPlugin";

export function DocumentEditor({ onChange }: { onChange: (editorState: EditorState) => void }) {
  return (
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
          <OnChangePlugin onChange={onChange} ignoreSelectionChange />
        </div>
      </div>
    </div>
  );
}
