import { lexicalJsonToAuthorTree } from "./authorTree";
import { authorTreeToSemanticGraph } from "./semanticGraph";
import { semanticGraphToOutputProjection } from "./outputProjection";
import type { LexicalEditorJson } from "./authorTree";

export function createPipelineFromLexicalJson(json: LexicalEditorJson) {
  const authorTree = lexicalJsonToAuthorTree(json);
  const semanticGraph = authorTreeToSemanticGraph(authorTree);
  const outputProjection = semanticGraphToOutputProjection(semanticGraph);

  return { authorTree, semanticGraph, outputProjection };
}
