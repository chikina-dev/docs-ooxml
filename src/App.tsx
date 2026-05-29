import { useCallback } from "react";
import { useAuthoringSession } from "./app/authoringSession";
import { DocumentWorkspace } from "./components/DocumentWorkspace";
import { PipelinePreview } from "./components/PipelinePreview";
import { downloadDocx } from "./effects/docxDownload";

export default function App() {
  const { commands, snapshot } = useAuthoringSession();

  const handleExport = useCallback(() => {
    downloadDocx({
      title: snapshot.title,
      projection: snapshot.pipeline.outputProjection,
    });
  }, [snapshot.pipeline.outputProjection, snapshot.title]);

  return (
    <main className="app-shell">
      <DocumentWorkspace commands={commands} onExport={handleExport} snapshot={snapshot} />
      <PipelinePreview pipeline={snapshot.pipeline} />
    </main>
  );
}
