import type { EditorState } from "lexical";
import { useCallback, useMemo, useState } from "react";
import { createPipelineFromLexicalJson } from "../pipeline/createPipeline.ts";
import type { AuthoringCommands, AuthoringSnapshot } from "./authoringTypes.ts";

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

export function useAuthoringSession(): {
  snapshot: AuthoringSnapshot;
  commands: AuthoringCommands;
} {
  const [title, setTitle] = useState("Untitled Document");
  const [pipeline, setPipeline] = useState(INITIAL_PIPELINE);

  const captureEditorState = useCallback((editorState: EditorState) => {
    setPipeline(createPipelineFromLexicalJson(editorState.toJSON()));
  }, []);

  const snapshot = useMemo(
    () => ({
      title,
      pipeline,
    }),
    [pipeline, title],
  );

  const commands = useMemo(
    () => ({
      setTitle,
      captureEditorState,
    }),
    [captureEditorState],
  );

  return { snapshot, commands };
}
