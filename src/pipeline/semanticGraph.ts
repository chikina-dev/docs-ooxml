import type {
  AuthorBlock,
  AuthorInline,
  AuthorTree,
  SemanticAuthorGraph,
  SemanticBlock,
  SemanticInline,
  SemanticSection,
} from "./types.ts";

type BuildState = {
  nextBlockId: number;
  nextSectionId: number;
  nextListGroupId: number;
  outlineCounters: number[];
  activeSections: Array<SemanticSection | undefined>;
  currentSectionId: string;
  sections: SemanticSection[];
  blocks: SemanticBlock[];
  listGroups: SemanticAuthorGraph["listGroups"];
  edges: SemanticAuthorGraph["edges"];
  readingOrder: string[];
};

export function authorTreeToSemanticGraph(tree: AuthorTree): SemanticAuthorGraph {
  const rootSection: SemanticSection = {
    id: "section-root",
    title: "Document",
    level: 0,
    outlinePath: [],
    childSectionIds: [],
    blockIds: [],
  };
  const state: BuildState = {
    nextBlockId: 1,
    nextSectionId: 1,
    nextListGroupId: 1,
    outlineCounters: [],
    activeSections: [rootSection],
    currentSectionId: rootSection.id,
    sections: [rootSection],
    blocks: [],
    listGroups: [],
    edges: [],
    readingOrder: [],
  };

  tree.children.forEach((block, index) => visitAuthorBlock(block, index, state));

  return {
    kind: "semanticAuthorGraph",
    rootSectionId: rootSection.id,
    readingOrder: state.readingOrder,
    sections: state.sections,
    listGroups: state.listGroups,
    blocks: state.blocks,
    edges: state.edges,
  };
}

function visitAuthorBlock(block: AuthorBlock, index: number, state: BuildState) {
  if (block.kind === "heading") {
    addHeading(block, index, state);
    return;
  }

  if (block.kind === "list") {
    addListGroup(block, index, state);
    return;
  }

  addParagraph(block, index, state);
}

function addHeading(
  block: Extract<AuthorBlock, { kind: "heading" }>,
  index: number,
  state: BuildState,
) {
  const parentSection = findParentSection(block.level, state);
  const headingBlockId = nextBlockId(state);
  const section = createSection(block, parentSection, headingBlockId, state);
  const semanticBlock: SemanticBlock = {
    id: headingBlockId,
    sourcePath: `/children/${index}`,
    sectionId: parentSection.id,
    createsSectionId: section.id,
    kind: "heading",
    level: block.level,
    title: plainText(block.children),
    outlinePath: section.outlinePath,
    inlines: normalizeInlines(block.children),
  };

  parentSection.childSectionIds.push(section.id);
  parentSection.blockIds.push(headingBlockId);
  state.sections.push(section);
  state.blocks.push(semanticBlock);
  state.readingOrder.push(headingBlockId);
  state.edges.push({ from: parentSection.id, to: headingBlockId, kind: "contains" });
  state.edges.push({ from: headingBlockId, to: section.id, kind: "opensSection" });
  state.activeSections[block.level] = section;
  state.activeSections.length = block.level + 1;
  state.currentSectionId = section.id;
}

function addParagraph(
  block: Extract<AuthorBlock, { kind: "paragraph" }>,
  index: number,
  state: BuildState,
) {
  const currentSection = sectionById(state.currentSectionId, state);
  const blockId = nextBlockId(state);
  state.blocks.push({
    id: blockId,
    sourcePath: `/children/${index}`,
    sectionId: currentSection.id,
    kind: "paragraph",
    inlines: normalizeInlines(block.children),
  });
  currentSection.blockIds.push(blockId);
  state.readingOrder.push(blockId);
  state.edges.push({ from: currentSection.id, to: blockId, kind: "contains" });
}

function addListGroup(
  block: Extract<AuthorBlock, { kind: "list" }>,
  index: number,
  state: BuildState,
) {
  const currentSection = sectionById(state.currentSectionId, state);
  const listGroup = {
    id: `list-${state.nextListGroupId++}`,
    sectionId: currentSection.id,
    listKind: block.listKind,
    itemIds: [] as string[],
  };

  block.children.forEach((item, itemIndex) => {
    const blockId = nextBlockId(state);
    listGroup.itemIds.push(blockId);
    state.blocks.push({
      id: blockId,
      sourcePath: `/children/${index}/children/${itemIndex}`,
      sectionId: currentSection.id,
      listGroupId: listGroup.id,
      ordinalInList: itemIndex + 1,
      kind: "listItem",
      listKind: block.listKind,
      level: 0,
      inlines: normalizeInlines(item.children),
    });
    currentSection.blockIds.push(blockId);
    state.readingOrder.push(blockId);
    state.edges.push({ from: currentSection.id, to: blockId, kind: "contains" });
    state.edges.push({ from: listGroup.id, to: blockId, kind: "continuesList" });
  });

  state.listGroups.push(listGroup);
}

function createSection(
  block: Extract<AuthorBlock, { kind: "heading" }>,
  parentSection: SemanticSection,
  headingBlockId: string,
  state: BuildState,
): SemanticSection {
  const outlinePath = nextOutlinePath(block.level, state);
  return {
    id: `section-${state.nextSectionId++}`,
    parentId: parentSection.id,
    headingBlockId,
    title: plainText(block.children),
    level: block.level,
    outlinePath,
    childSectionIds: [],
    blockIds: [],
  };
}

function findParentSection(level: 1 | 2 | 3, state: BuildState): SemanticSection {
  for (let candidateLevel = level - 1; candidateLevel >= 0; candidateLevel -= 1) {
    const section = state.activeSections[candidateLevel];
    if (section) {
      return section;
    }
  }

  return state.sections[0]!;
}

function nextOutlinePath(level: 1 | 2 | 3, state: BuildState): number[] {
  state.outlineCounters[level - 1] = (state.outlineCounters[level - 1] ?? 0) + 1;
  state.outlineCounters.length = level;

  return state.outlineCounters.slice(0, level);
}

function nextBlockId(state: BuildState): string {
  return `block-${state.nextBlockId++}`;
}

function sectionById(sectionId: string, state: BuildState): SemanticSection {
  return state.sections.find((section) => section.id === sectionId) ?? state.sections[0]!;
}

function normalizeInlines(inlines: AuthorInline[]): SemanticInline[] {
  const normalized: SemanticInline[] = [];

  for (const inline of inlines) {
    if (inline.kind === "lineBreak") {
      normalized.push({ kind: "break" });
      continue;
    }

    const previous = normalized.at(-1);
    if (
      previous?.kind === "span" &&
      previous.marks.bold === inline.marks.bold &&
      previous.marks.italic === inline.marks.italic &&
      previous.marks.underline === inline.marks.underline
    ) {
      previous.text += inline.text;
      continue;
    }

    normalized.push({ kind: "span", text: inline.text, marks: { ...inline.marks } });
  }

  return normalized;
}

function plainText(inlines: AuthorInline[]): string {
  return inlines.map((inline) => (inline.kind === "textRun" ? inline.text : "\n")).join("");
}
