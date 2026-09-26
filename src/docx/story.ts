import { OoxmlError } from "../errors.ts";
import { getContentType, OpcPackage } from "../opc/index.ts";
import { attribute, parseXml, type XmlElement } from "../xml/index.ts";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING_NAMESPACES = new Set<string>([
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  "urn:schemas-microsoft-com:vml",
  "urn:schemas-microsoft-com:office:office",
  "urn:schemas-microsoft-com:office:word",
]);
const KNOWN_NAMESPACES = new Set<string>([W_NS, MC_NS, R_NS, ...DRAWING_NAMESPACES]);
const STORY_RELATIONSHIP_KIND = new Map<string, StoryKind>([
  [`${OFFICE_REL}header`, "header"],
  [`${OFFICE_REL}footer`, "footer"],
  [`${OFFICE_REL}footnotes`, "footnotes"],
  [`${OFFICE_REL}endnotes`, "endnotes"],
  [`${OFFICE_REL}comments`, "comments"],
]);
const STORY_CONTENT_TYPE_KIND = new Map<string, StoryKind>([
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "body"],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", "header"],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", "footer"],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", "footnotes"],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml", "endnotes"],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", "comments"],
]);
const STORY_KIND_ORDER = new Map<StoryKind, number>([
  ["body", 0],
  ["header", 1],
  ["footer", 2],
  ["footnotes", 3],
  ["endnotes", 4],
  ["comments", 5],
]);
const BLIND_KIND_ORDER = new Map<string, number>([
  ["alternate-content", 0],
  ["field", 1],
  ["textbox", 2],
  ["drawing", 3],
  ["unknown-namespace", 4],
  ["external-story-relationship", 5],
  ["orphan-story-part", 6],
  ["unsupported-story-part", 7],
]);
const BLOCK_INSERTION_TAGS = new Set(["ins", "moveTo"]);
const BLOCK_DELETION_TAGS = new Set(["del", "moveFrom"]);

export type RevisionView = "current" | "original" | "all";
export type StoryKind = "body" | "header" | "footer" | "footnotes" | "endnotes" | "comments";

export type Story = {
  part: string;
  kind: StoryKind;
  paragraphs: Array<{
    index: number;
    text: string;
  }>;
};

export type StoryInspection = {
  stories: Story[];
  blindRegions: Array<{
    part: string;
    kind: string;
    count: number;
  }>;
};

type StoryPart = {
  part: string;
  kind: StoryKind;
  discoveryOrder: number;
};

type StoryDiscovery = {
  stories: StoryPart[];
  blindRegions: Array<{ part: string; kind: string; count: number }>;
  visited: Set<string>;
};

type TextState = {
  view: RevisionView;
  fieldStack: boolean[];
};

type TextSegment = {
  text: string;
  source: "normal" | "inserted" | "deleted";
};

export function storyParts(pkg: OpcPackage): Array<{ part: string; kind: StoryKind }> {
  return discoverStoryParts(pkg).stories.map(({ part, kind }) => ({ part, kind }));
}

export function inspectStories(
  pkg: OpcPackage,
  options: { view?: RevisionView } = {},
): StoryInspection {
  const view = options.view ?? "current";
  if (view !== "current" && view !== "original" && view !== "all") {
    throw new OoxmlError("docx-story-invalid-view", `Unsupported DOCX story view: ${String(view)}`);
  }

  const discovery = discoverStoryParts(pkg);
  const blindCounts = new Map<string, number>();
  addBlindRegions(blindCounts, discovery.blindRegions);

  const stories: Story[] = [];
  for (const story of discovery.stories) {
    const document = parseXml(pkg.text(story.part));
    addBlindRegions(blindCounts, countBlindRegions(story.part, document.root));

    const containers = storyContainers(story, document.root);
    if (!containers) {
      incrementBlind(blindCounts, story.part, "unsupported-story-part");
      continue;
    }

    const paragraphs: Story["paragraphs"] = [];
    for (const container of containers) {
      collectContainerBlocks(container, view, paragraphs);
    }
    stories.push({ part: story.part, kind: story.kind, paragraphs });
  }

  for (const name of pkg.names()) {
    if (discovery.visited.has(name) || name === "[Content_Types].xml" || name.endsWith(".rels")) {
      continue;
    }
    const kind = STORY_CONTENT_TYPE_KIND.get(getContentType(pkg, name) ?? "");
    if (kind && kind !== "body") {
      incrementBlind(blindCounts, name, "orphan-story-part");
    }
  }

  return {
    stories,
    blindRegions: [...blindCounts.entries()]
      .map(([key, count]) => {
        const separator = key.indexOf("\u0000");
        return {
          part: key.slice(0, separator),
          kind: key.slice(separator + 1),
          count,
        };
      })
      .sort((left, right) =>
        left.part.localeCompare(right.part) ||
        blindKindSortValue(left.kind) - blindKindSortValue(right.kind) ||
        left.kind.localeCompare(right.kind),
      ),
  };
}

function discoverStoryParts(pkg: OpcPackage): StoryDiscovery {
  const main = pkg.mainPart();
  const visited = new Set<string>([main]);
  const queue = [main];
  const discovered: StoryPart[] = [{ part: main, kind: "body", discoveryOrder: -1 }];
  const blindCounts = new Map<string, number>();
  let discoveryOrder = 0;

  for (let index = 0; index < queue.length; index += 1) {
    const owner = queue[index]!;
    for (const relationship of pkg.relationships(owner)) {
      const kind = STORY_RELATIONSHIP_KIND.get(relationship.type);
      if (!kind) {
        continue;
      }
      if (relationship.external) {
        incrementBlind(blindCounts, owner, "external-story-relationship");
        continue;
      }
      if (!relationship.resolved || visited.has(relationship.resolved)) {
        continue;
      }
      visited.add(relationship.resolved);
      queue.push(relationship.resolved);
      discovered.push({ part: relationship.resolved, kind, discoveryOrder });
      discoveryOrder += 1;
    }
  }

  const stories = [
    discovered[0]!,
    ...discovered
      .slice(1)
      .sort((left, right) =>
        storyKindSortValue(left.kind) - storyKindSortValue(right.kind) ||
        left.discoveryOrder - right.discoveryOrder ||
        left.part.localeCompare(right.part),
      ),
  ];

  return {
    stories,
    visited,
    blindRegions: [...blindCounts.entries()].map(([key, count]) => {
      const separator = key.indexOf("\u0000");
      return {
        part: key.slice(0, separator),
        kind: key.slice(separator + 1),
        count,
      };
    }),
  };
}

function storyContainers(story: { part: string; kind: StoryKind }, root: XmlElement): XmlElement[] | undefined {
  if (root.namespaceURI !== W_NS) {
    return undefined;
  }

  if (story.kind === "body") {
    if (root.localName !== "document") {
      return undefined;
    }
    return root.children.filter((child) => child.namespaceURI === W_NS && child.localName === "body");
  }

  if (story.kind === "header") {
    return root.localName === "hdr" ? [root] : undefined;
  }

  if (story.kind === "footer") {
    return root.localName === "ftr" ? [root] : undefined;
  }

  if (story.kind === "footnotes" || story.kind === "endnotes") {
    const expectedRoot = story.kind === "footnotes" ? "footnotes" : "endnotes";
    const expectedChild = story.kind === "footnotes" ? "footnote" : "endnote";
    if (root.localName !== expectedRoot) {
      return undefined;
    }
    return root.children.filter((child) => {
      if (child.namespaceURI !== W_NS || child.localName !== expectedChild) {
        return false;
      }
      const noteType = attribute(child, "type", W_NS);
      return noteType !== "separator" && noteType !== "continuationSeparator";
    });
  }

  if (story.kind === "comments") {
    if (root.localName !== "comments") {
      return undefined;
    }
    return root.children.filter((child) => child.namespaceURI === W_NS && child.localName === "comment");
  }

  return undefined;
}

function collectContainerBlocks(
  container: XmlElement,
  view: RevisionView,
  paragraphs: Story["paragraphs"],
): void {
  for (const child of selectedChildren(container)) {
    if (child.namespaceURI === W_NS) {
      if (BLOCK_INSERTION_TAGS.has(child.localName)) {
        if (view !== "original") {
          collectContainerBlocks(child, view, paragraphs);
        }
        continue;
      }
      if (BLOCK_DELETION_TAGS.has(child.localName)) {
        if (view !== "current") {
          collectContainerBlocks(child, view, paragraphs);
        }
        continue;
      }
      if (child.localName === "sdt") {
        const content = child.children.find((candidate) => candidate.namespaceURI === W_NS && candidate.localName === "sdtContent");
        if (content) {
          collectContainerBlocks(content, view, paragraphs);
        }
        continue;
      }
      if (child.localName === "p") {
        paragraphs.push({ index: paragraphs.length, text: collectParagraphText(child, view) });
        continue;
      }
      if (child.localName === "tbl") {
        paragraphs.push({ index: paragraphs.length, text: collectTableText(child, view) });
        continue;
      }
    }

    collectContainerBlocks(child, view, paragraphs);
  }
}

function collectParagraphText(paragraph: XmlElement, view: RevisionView): string {
  const segments: TextSegment[] = [];
  const state: TextState = { view, fieldStack: [] };
  for (const child of selectedChildren(paragraph)) {
    collectText(child, state, segments, "normal");
  }
  return materializeText(segments, view);
}

function collectTableText(table: XmlElement, view: RevisionView): string {
  const rowTexts: string[] = [];
  for (const row of table.children) {
    if (row.namespaceURI !== W_NS || row.localName !== "tr") {
      continue;
    }
    const cellTexts: string[] = [];
    for (const cell of row.children) {
      if (cell.namespaceURI !== W_NS || cell.localName !== "tc") {
        continue;
      }
      cellTexts.push(collectCellText(cell, view));
    }
    rowTexts.push(cellTexts.join("\t"));
  }
  return rowTexts.join("\n");
}

function collectCellText(cell: XmlElement, view: RevisionView): string {
  const blocks: Story["paragraphs"] = [];
  collectContainerBlocks(cell, view, blocks);
  return blocks.map((block) => block.text).join("\n");
}

function collectText(
  element: XmlElement,
  state: TextState,
  segments: TextSegment[],
  source: TextSegment["source"],
): void {
  if (element.namespaceURI === W_NS) {
    if (element.localName === "txbxContent") {
      return;
    }
    if (BLOCK_INSERTION_TAGS.has(element.localName)) {
      if (state.view !== "original") {
        const nextSource = state.view === "all" ? "inserted" : source;
        for (const child of selectedChildren(element)) {
          collectText(child, state, segments, nextSource);
        }
      }
      return;
    }
    if (BLOCK_DELETION_TAGS.has(element.localName)) {
      if (state.view !== "current") {
        const nextSource = state.view === "all" ? "deleted" : source;
        for (const child of selectedChildren(element)) {
          collectText(child, state, segments, nextSource);
        }
      }
      return;
    }
    if (element.localName === "instrText" || element.localName === "delInstrText") {
      return;
    }
    if (element.localName === "fldChar") {
      const fieldType = attribute(element, "fldCharType", W_NS);
      if (fieldType === "begin") {
        state.fieldStack.push(false);
      } else if (fieldType === "separate") {
        const lastIndex = state.fieldStack.length - 1;
        if (lastIndex >= 0) {
          state.fieldStack[lastIndex] = true;
        }
      } else if (fieldType === "end") {
        state.fieldStack.pop();
      }
      return;
    }
    if (element.localName === "t") {
      if (fieldTextVisible(state)) {
        pushText(segments, element.text, source);
      }
      return;
    }
    if (element.localName === "delText") {
      if (state.view !== "current" && fieldTextVisible(state)) {
        pushText(segments, element.text, source);
      }
      return;
    }
    if (element.localName === "tab") {
      if (fieldTextVisible(state)) {
        pushText(segments, "\t", source);
      }
      return;
    }
    if (element.localName === "br" || element.localName === "cr") {
      if (fieldTextVisible(state)) {
        pushText(segments, "\n", source);
      }
      return;
    }
    if (element.localName === "noBreakHyphen") {
      if (fieldTextVisible(state)) {
        pushText(segments, "-", source);
      }
      return;
    }
    if (element.localName === "softHyphen") {
      if (fieldTextVisible(state)) {
        pushText(segments, "\u00ad", source);
      }
      return;
    }
  }

  for (const child of selectedChildren(element)) {
    collectText(child, state, segments, source);
  }
}

function fieldTextVisible(state: TextState): boolean {
  return state.fieldStack.every((showResult) => showResult);
}

function pushText(
  segments: TextSegment[],
  text: string,
  source: TextSegment["source"],
): void {
  if (text.length === 0) {
    return;
  }
  const last = segments.at(-1);
  if (last && last.source === source) {
    last.text += text;
    return;
  }
  segments.push({ text, source });
}

function materializeText(segments: TextSegment[], view: RevisionView): string {
  if (view !== "all") {
    return segments.map((segment) => segment.text).join("");
  }

  const merged: TextSegment[] = [];
  for (let index = 0; index < segments.length; index += 1) {
    const current = segments[index]!;
    const next = segments[index + 1];
    if (
      next &&
      current.source !== "normal" &&
      next.source !== "normal" &&
      current.source !== next.source
    ) {
      pushText(merged, mergeAlternativeText(current.text, next.text), "normal");
      index += 1;
      continue;
    }
    pushText(merged, current.text, current.source);
  }
  return merged.map((segment) => segment.text).join("");
}

function mergeAlternativeText(left: string, right: string): string {
  const prefix = commonPrefixLength(left, right);
  const suffix = commonSuffixLength(left.slice(prefix), right.slice(prefix));
  return (
    left.slice(0, prefix) +
    left.slice(prefix, left.length - suffix) +
    right.slice(prefix, right.length - suffix) +
    left.slice(left.length - suffix)
  );
}

function commonPrefixLength(left: string, right: string): number {
  const limit = Math.min(left.length, right.length);
  let index = 0;
  while (index < limit && left[index] === right[index]) {
    index += 1;
  }
  return index;
}

function commonSuffixLength(left: string, right: string): number {
  const limit = Math.min(left.length, right.length);
  let index = 0;
  while (index < limit && left[left.length - 1 - index] === right[right.length - 1 - index]) {
    index += 1;
  }
  return index;
}

function countBlindRegions(part: string, root: XmlElement): Array<{ part: string; kind: string; count: number }> {
  const counts = new Map<string, number>();

  const walk = (element: XmlElement): void => {
    for (const child of element.children) {
      if (child.namespaceURI === MC_NS && child.localName === "AlternateContent") {
        incrementBlind(counts, part, "alternate-content");
        const branch = selectAlternateContentBranch(child);
        if (branch) {
          for (const selected of selectedChildren(branch)) {
            walk(selected);
          }
        }
        continue;
      }

      if (child.namespaceURI === W_NS) {
        if (child.localName === "fldSimple") {
          incrementBlind(counts, part, "field");
        } else if (child.localName === "fldChar" && attribute(child, "fldCharType", W_NS) === "begin") {
          incrementBlind(counts, part, "field");
        } else if (child.localName === "txbxContent") {
          incrementBlind(counts, part, "textbox");
          continue;
        } else if (child.localName === "drawing" || child.localName === "pict" || child.localName === "object") {
          incrementBlind(counts, part, "drawing");
        }
      } else if (!KNOWN_NAMESPACES.has(child.namespaceURI)) {
        incrementBlind(counts, part, "unknown-namespace");
      }

      walk(child);
    }
  };

  walk(root);
  return [...counts.entries()].map(([key, count]) => {
    const separator = key.indexOf("\u0000");
    return {
      part: key.slice(0, separator),
      kind: key.slice(separator + 1),
      count,
    };
  });
}

function selectedChildren(element: XmlElement): XmlElement[] {
  const selected: XmlElement[] = [];
  for (const child of element.children) {
    if (child.namespaceURI === MC_NS && child.localName === "AlternateContent") {
      const branch = selectAlternateContentBranch(child);
      if (branch) {
        selected.push(...selectedChildren(branch));
      }
      continue;
    }
    selected.push(child);
  }
  return selected;
}

function selectAlternateContentBranch(element: XmlElement): XmlElement | undefined {
  for (const child of element.children) {
    if (child.namespaceURI === MC_NS && child.localName === "Choice") {
      return child;
    }
  }
  for (const child of element.children) {
    if (child.namespaceURI === MC_NS && child.localName === "Fallback") {
      return child;
    }
  }
  return undefined;
}

function addBlindRegions(
  counts: Map<string, number>,
  regions: Array<{ part: string; kind: string; count: number }>,
): void {
  for (const region of regions) {
    const key = blindKey(region.part, region.kind);
    counts.set(key, (counts.get(key) ?? 0) + region.count);
  }
}

function incrementBlind(counts: Map<string, number>, part: string, kind: string): void {
  const key = blindKey(part, kind);
  counts.set(key, (counts.get(key) ?? 0) + 1);
}

function blindKey(part: string, kind: string): string {
  return `${part}\u0000${kind}`;
}

function storyKindSortValue(kind: StoryKind): number {
  return STORY_KIND_ORDER.get(kind) ?? Number.MAX_SAFE_INTEGER;
}

function blindKindSortValue(kind: string): number {
  return BLIND_KIND_ORDER.get(kind) ?? Number.MAX_SAFE_INTEGER;
}
