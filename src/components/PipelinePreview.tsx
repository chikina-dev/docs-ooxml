import type { PipelineSnapshot } from "../app/authoringTypes";

type PreviewData =
  | PipelineSnapshot["authorTree"]
  | PipelineSnapshot["semanticGraph"]
  | PipelineSnapshot["outputProjection"];

export function PipelinePreview({ pipeline }: { pipeline: PipelineSnapshot }) {
  return (
    <aside className="pipeline-panel" aria-label="Pipeline preview">
      <PreviewSection title="Author Tree" data={pipeline.authorTree} />
      <PreviewSection title="Semantic Author Graph" data={pipeline.semanticGraph} />
      <PreviewSection title="Output Projection" data={pipeline.outputProjection} />
    </aside>
  );
}

function PreviewSection({ title, data }: { title: string; data: PreviewData }) {
  return (
    <details className="preview-section" open>
      <summary>{title}</summary>
      <pre>{JSON.stringify(data, null, 2)}</pre>
    </details>
  );
}
