import { useCallback } from "react";
import { createExportDocxCommand } from "./app/exportCommand";
import { useAuthoringSession } from "./app/authoringSession";
import { DocumentWorkspace } from "./components/DocumentWorkspace";
import { PipelinePreview } from "./components/PipelinePreview";
import { downloadDocx } from "./effects/docxDownload";

export default function App() {
  const { commands, snapshot } = useAuthoringSession();

  const handleExport = useCallback(() => {
    downloadDocx(createExportDocxCommand(snapshot));
  }, [snapshot]);

  return (
    <main className="app-shell">
      <DocumentWorkspace commands={commands} onExport={handleExport} snapshot={snapshot} />
      <PipelinePreview pipeline={snapshot.pipeline} />
    </main>
  );
}
