import type { EditorState } from "lexical";
import type { createPipelineFromLexicalJson } from "../pipeline/createPipeline.ts";

export type PipelineSnapshot = ReturnType<typeof createPipelineFromLexicalJson>;

export type AuthoringSnapshot = {
  title: string;
  pipeline: PipelineSnapshot;
};

export type AuthoringCommands = {
  setTitle: (title: string) => void;
  captureEditorState: (editorState: EditorState) => void;
};
