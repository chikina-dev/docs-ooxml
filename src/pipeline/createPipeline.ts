import { lexicalJsonToAuthorTree } from "./authorTree.ts";
import { authorTreeToSemanticGraph } from "./semanticGraph.ts";
import { semanticGraphToOutputProjection } from "./outputProjection.ts";

export function createPipelineFromLexicalJson(json: unknown) {
  const authorTree = lexicalJsonToAuthorTree(json as Parameters<typeof lexicalJsonToAuthorTree>[0]);
  const semanticGraph = authorTreeToSemanticGraph(authorTree);
  const outputProjection = semanticGraphToOutputProjection(semanticGraph);

  return { authorTree, semanticGraph, outputProjection };
}
