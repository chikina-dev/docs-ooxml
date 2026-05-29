import { LexicalComposer } from "@lexical/react/LexicalComposer";
import type { AuthoringCommands, AuthoringSnapshot } from "../app/authoringTypes";
import { lexicalInitialConfig } from "../editor/lexicalConfig";
import { DocumentEditor } from "./DocumentEditor";
import { EditorToolbar } from "./EditorToolbar";
import { TitleBar } from "./TitleBar";

export function DocumentWorkspace({
  commands,
  onExport,
  snapshot,
}: {
  commands: AuthoringCommands;
  onExport: () => void;
  snapshot: AuthoringSnapshot;
}) {
  return (
    <section className="workspace">
      <TitleBar commands={commands} snapshot={snapshot} />
      <LexicalComposer initialConfig={lexicalInitialConfig}>
        <EditorToolbar onExport={onExport} />
        <DocumentEditor onChange={commands.captureEditorState} />
      </LexicalComposer>
    </section>
  );
}
