import { useCallback } from "react";
import { useAuthoringSession } from "./app/authoringSession.ts";
import { DocumentWorkspace } from "./components/DocumentWorkspace.tsx";
import { PipelinePreview } from "./components/PipelinePreview.tsx";
import { downloadDocx } from "./effects/docxDownload.ts";
import type { DocxWriteStrategy } from "./ooxml/docx.ts";

export default function App() {
  const { commands, snapshot } = useAuthoringSession();

  const handleExport = useCallback(
    (strategy: DocxWriteStrategy) => {
      downloadDocx({
        title: snapshot.title,
        projection: snapshot.pipeline.outputProjection,
        strategy,
      });
    },
    [snapshot.pipeline.outputProjection, snapshot.title],
  );

  return (
    <main className="app-shell">
      <DocumentWorkspace commands={commands} onExport={handleExport} snapshot={snapshot} />
      <PipelinePreview pipeline={snapshot.pipeline} />
    </main>
  );
}
