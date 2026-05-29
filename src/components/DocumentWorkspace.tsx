import { LexicalComposer } from "@lexical/react/LexicalComposer";
import type { AuthoringCommands, AuthoringSnapshot } from "../app/authoringTypes.ts";
import { lexicalInitialConfig } from "../editor/lexicalConfig.ts";
import type { DocxWriteStrategy } from "../ooxml/docx.ts";
import { DocumentEditor } from "./DocumentEditor.tsx";
import { EditorToolbar } from "./EditorToolbar.tsx";
import { TitleBar } from "./TitleBar.tsx";

export function DocumentWorkspace({
  commands,
  onExport,
  snapshot,
}: {
  commands: AuthoringCommands;
  onExport: (strategy: DocxWriteStrategy) => void;
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
