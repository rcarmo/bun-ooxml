import { OoxmlError } from "../errors.ts";
import {runPropertyChangeProblem,restoredRunProperties} from './revision-properties.ts';
import {scanRunMoves} from './revision-moves.ts';
import { OpcPackage } from "../opc/index.ts";
import { applyEdits, attribute, parseXml, type XmlDocument, type XmlElement } from "../xml/index.ts";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

const STORY_KIND_BY_SUFFIX = {
  header: "header",
  footer: "footer",
  footnotes: "footnotes",
  endnotes: "endnotes",
  comments: "comments",
} as const;

const STORY_ORDER = ["body", "header", "footer", "footnotes", "endnotes", "comments"] as const;
const MOVE_TAGS = new Set(["moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"]);
const FORMAT_TAGS = new Set(["rPrChange", "pPrChange", "sectPrChange", "numberingChange"]);
const TABLE_TAGS = new Set(["tblPrChange", "tblPrExChange", "tblGridChange", "trPrChange", "tcPrChange", "cellIns", "cellDel", "cellMerge"]);
const CUSTOM_XML_TAGS = new Set(["customXmlInsRangeStart", "customXmlDelRangeStart", "customXmlMoveFromRangeStart", "customXmlMoveToRangeStart"]);
const REVISION_CONTAINER_TAGS = new Set(["ins", "del", ...MOVE_TAGS, ...FORMAT_TAGS, ...TABLE_TAGS, ...CUSTOM_XML_TAGS]);

type StoryKind = typeof STORY_ORDER[number];

type StoryPart = {
  part: string;
  kind: StoryKind;
};

type SupportedRevisionEntry = Revision & {
  element: XmlElement;
};

type ScannedStoryPart = {
  part: string;
  kind: StoryKind;
  xml: string;
  document: XmlDocument;
  revisions: SupportedRevisionEntry[];
  unsupported: RevisionFinding[];
  moveMarkers: XmlElement[];
};

export type Revision = {
  part: string;
  id: string;
  kind: "insertion" | "deletion" | "run-properties" | "move-from" | "move-to";
  author?: string;
  date?: string;
  text: string;
};

export type RevisionFinding = {
  part: string;
  kind: string;
  reason: string;
};

export type RevisionProfile = 'text-only' | 'text-and-run-properties' | 'text-properties-and-moves';
function revisionProfile(value:unknown):RevisionProfile {
  if(value===undefined)return 'text-only';
  if(value!=='text-only'&&value!=='text-and-run-properties'&&value!=='text-properties-and-moves')throw new OoxmlError('docx-revisions-invalid-profile','Unknown revision profile');
  return value;
}
export function inspectRevisions(pkg: OpcPackage, options:{profile?:RevisionProfile}={}): { revisions: Revision[]; unsupported: RevisionFinding[] } {
  const profile=revisionProfile(options.profile);
  const scanned = discoverStoryParts(pkg).map((story) => scanStoryPart(pkg, story,profile));
  return {
    revisions: scanned.flatMap((story) => story.revisions.map(({ element: _element, ...revision }) => revision)),
    unsupported: scanned.flatMap((story) => story.unsupported),
  };
}

export function resolveRevisions(
  pkg: OpcPackage,
  action: "accept" | "reject",
  options: { parts?: string[]; profile?:RevisionProfile } = {},
): { resolved: number; changedParts: string[] } {
  if (action !== "accept" && action !== "reject") {
    throw new OoxmlError("docx-revisions-invalid-action", `Unsupported revision action: ${String(action)}`);
  }

  const profile=revisionProfile(options.profile);
  const stories = discoverStoryParts(pkg);
  const selected = selectStoryParts(stories, options.parts);
  if (selected.length === 0) {
    return { resolved: 0, changedParts: [] };
  }

  const scanned = selected.map((story) => scanStoryPart(pkg, story,profile));
  const unsupported = scanned.flatMap((story) => [
    ...story.unsupported,
    ...namespaceLiftFindings(story, action),
  ]);
  if (unsupported.length > 0) {
    throw new OoxmlError("docx-revisions-unsupported", describeUnsupported(unsupported));
  }

  if(profile!=='text-only'&&isProtected(pkg,true))throw new OoxmlError('docx-revisions-protected','Cannot resolve run-property revisions with protected or external settings');
  const resolved = scanned.reduce((sum, story) => sum + story.revisions.length, 0);
  if (resolved === 0) {
    return { resolved: 0, changedParts: [] };
  }

  if (isProtected(pkg)) {
    throw new OoxmlError(
      "docx-revisions-protected",
      "Cannot resolve revisions while word/settings.xml enforces Restrict Editing (w:documentProtection)",
    );
  }

  const plans = scanned
    .map((story) => ({ story, edits: buildStoryEdits(story, action) }))
    .filter((plan) => plan.edits.length > 0);

  if (plans.length === 0) {
    return { resolved: 0, changedParts: [] };
  }

  pkg.transaction(() => {
    for (const plan of plans) {
      const updated = applyEdits(plan.story.xml, plan.edits);
      if (updated !== plan.story.xml) {
        const bytes=pkg.get(plan.story.part)!;
        pkg.set(plan.story.part,bytes[0]===239&&bytes[1]===187&&bytes[2]===191?'\ufeff'+updated:updated);
        parseXml(pkg.text(plan.story.part));
      }
    }
    if(profile!=='text-only')for(const plan of plans){
      const result=scanStoryPart(pkg,{part:plan.story.part,kind:plan.story.kind},profile);
      if(result.revisions.length||result.unsupported.length)throw new OoxmlError('docx-revisions-validation','Resolution left revision markup or unsupported content');
    }
    pkg.toBytes();
  });

  return {
    resolved,
    changedParts: plans.map((plan) => plan.story.part),
  };
}

function discoverStoryParts(pkg: OpcPackage): StoryPart[] {
  const main = pkg.mainPart();
  const grouped = new Map<StoryKind, Array<{ order: number; part: string }>>();
  const seen = new Set<string>([main]);
  const pending = [main];
  let nextOwner = 0;
  let discoveryOrder = 0;

  while (nextOwner < pending.length) {
    const owner = pending[nextOwner++]!;
    for (const relationship of pkg.relationships(owner)) {
      if (relationship.external || !relationship.resolved || !relationship.type.startsWith(OFFICE_REL)) {
        continue;
      }

      const target = relationship.resolved;
      if (seen.has(target)) {
        continue;
      }
      seen.add(target);
      pending.push(target);

      const suffix = relationship.type.slice(OFFICE_REL.length);
      const kind = STORY_KIND_BY_SUFFIX[suffix as keyof typeof STORY_KIND_BY_SUFFIX];
      if (!kind) {
        continue;
      }

      const items = grouped.get(kind) ?? [];
      items.push({ order: discoveryOrder++, part: target });
      grouped.set(kind, items);
    }
  }

  const parts: StoryPart[] = [{ part: main, kind: "body" }];
  for (const kind of STORY_ORDER.slice(1)) {
    const items = grouped.get(kind);
    if (!items) {
      continue;
    }
    items
      .sort((left, right) => left.order - right.order || left.part.localeCompare(right.part))
      .forEach((item) => parts.push({ part: item.part, kind }));
  }
  return parts;
}

function selectStoryParts(stories: StoryPart[], requested?: string[]): StoryPart[] {
  if (!requested) {
    return stories;
  }

  const known = new Set(stories.map((story) => story.part));
  const requestedSet = new Set(requested);
  const unknown = [...requestedSet].filter((part) => !known.has(part)).sort();
  if (unknown.length > 0) {
    throw new OoxmlError(
      "docx-revisions-unknown-part",
      `Unknown DOCX story part selection: ${unknown.join(", ")}`,
    );
  }

  return stories.filter((story) => requestedSet.has(story.part));
}

function scanStoryPart(pkg: OpcPackage, story: StoryPart, profile:RevisionProfile='text-only'): ScannedStoryPart {
  const xml = pkg.text(story.part);
  const document = parseXml(xml);
  const revisions: SupportedRevisionEntry[] = [];
  const unsupported: RevisionFinding[] = [];
  const unsupportedKeys = new Set<string>();
  const ids = new Map<string, number>();
  const moves=profile==='text-properties-and-moves'?scanRunMoves(document.elements,xml):undefined;
  for(const reason of moves?.problems??[])pushFinding(unsupported,unsupportedKeys,story.part,'move',reason);
  for(const entry of moves?.moves??[])revisions.push({part:story.part,id:attribute(entry.element,'id',W_NS)!,kind:entry.kind,author:attribute(entry.element,'author',W_NS),date:attribute(entry.element,'date',W_NS),text:entry.text,element:entry.element});
  for(const element of moves?.identifiers??[]){const id=String(Number(attribute(element,'id',W_NS)));ids.set(id,(ids.get(id)??0)+1);}

  for (const element of document.elements) {
    if (element.namespaceURI !== W_NS) {
      continue;
    }

    const propertyChange=profile!=='text-only'&&element.localName==='rPrChange';
    if (element.localName === "ins" || element.localName === "del" || propertyChange) {
      const problem=propertyChange?runPropertyChangeProblem(element,xml):undefined;
      const support = propertyChange?(problem?{supported:false as const,kind:'format',reason:problem}:{supported:true as const}):inspectRunRevision(element,xml);
      if (!support.supported) {
        pushFinding(unsupported, unsupportedKeys, story.part, support.kind, support.reason, element.start);
        continue;
      }

      const id = attribute(element, "id", W_NS);
      if (!id || !/^[0-9]+$/.test(id) || !Number.isSafeInteger(Number(id))) {
        pushFinding(unsupported, unsupportedKeys, story.part, "identifier", "revision w:id must be a nonnegative safe decimal integer", element.start);
        continue;
      }

      const entry: SupportedRevisionEntry = {
        part: story.part,
        id,
        kind: propertyChange?'run-properties':element.localName === "ins" ? "insertion" : "deletion",
        author: attribute(element, "author", W_NS),
        date: attribute(element, "date", W_NS),
        text: revisionText(propertyChange?element.parent!.parent!:element),
        element,
      };
      revisions.push(entry);
      const numericId=String(Number(id));
      ids.set(numericId, (ids.get(numericId) ?? 0) + 1);
      continue;
    }

    if (MOVE_TAGS.has(element.localName)) {
      if(moves)continue;
      pushFinding(unsupported, unsupportedKeys, story.part, "move", "move revisions are not supported", element.start);
      continue;
    }
    if (FORMAT_TAGS.has(element.localName)) {
      pushFinding(unsupported, unsupportedKeys, story.part, "format", "format revisions are not supported", element.start);
      continue;
    }
    if (TABLE_TAGS.has(element.localName)) {
      pushFinding(unsupported, unsupportedKeys, story.part, "table", "table revisions are not supported", element.start);
      continue;
    }
    if (CUSTOM_XML_TAGS.has(element.localName)) {
      pushFinding(unsupported, unsupportedKeys, story.part, "customXml", "custom XML revisions are not supported", element.start);
    }
  }

  for (const [id, count] of ids) {
    if (count > 1) {
      pushFinding(
        unsupported,
        unsupportedKeys,
        story.part,
        "identifier",
        `duplicate revision id ${JSON.stringify(id)} in ${story.part}`,
      );
    }
  }

  revisions.sort((a,b)=>a.element.start-b.element.start);
  return { part: story.part, kind: story.kind, xml, document, revisions, unsupported, moveMarkers:moves?.markers??[] };
}

function inspectRunRevision(element: XmlElement, xml:string):
  | { supported: true }
  | { supported: false; kind: string; reason: string } {
  const parent = element.parent;
  if (!parent || parent.namespaceURI !== W_NS || parent.localName !== "p") {
    return unsupportedRevisionContext(element);
  }

  if (hasRevisionAncestor(element)) {
    return { supported: false, kind: "nested", reason: "nested revisions are not supported" };
  }

  for (const descendant of traverseDescendants(element)) {
    if (descendant.namespaceURI !== W_NS || !REVISION_CONTAINER_TAGS.has(descendant.localName)) {
      continue;
    }
    if (descendant.localName === "ins" || descendant.localName === "del") {
      return { supported: false, kind: "nested", reason: "nested revisions are not supported" };
    }
    if (MOVE_TAGS.has(descendant.localName)) {
      return { supported: false, kind: "move", reason: "move revisions are not supported" };
    }
    if (FORMAT_TAGS.has(descendant.localName)) {
      return { supported: false, kind: "format", reason: "format revisions are not supported" };
    }
    if (TABLE_TAGS.has(descendant.localName)) {
      return { supported: false, kind: "table", reason: "table revisions are not supported" };
    }
    return { supported: false, kind: "content", reason: "unsupported revision content was found inside the wrapper" };
  }

  if(hasNonElementContent(element,xml))return {supported:false,kind:'content',reason:'Revision wrapper contains unowned mixed content'};
  for (const child of element.children) {
    if(hasNonElementContent(child,xml))return {supported:false,kind:'content',reason:'Revision run contains unowned mixed content'};
    if (child.namespaceURI !== W_NS || child.localName !== "r") {
      return { supported: false, kind: "content", reason: "revision wrappers must contain only plain w:r children" };
    }
    const run = inspectRevisionRun(child, element.localName === "ins" ? "t" : "delText");
    if (!run.supported) {
      return run;
    }
  }

  return { supported: true };
}

function hasNonElementContent(element:XmlElement,xml:string):boolean {
  if(element.selfClosing)return false;
  let offset=element.openEnd;
  for(const child of element.children){if(!/^[ \t\r\n]*$/.test(xml.slice(offset,child.start)))return true;offset=child.end;}
  return !/^[ \t\r\n]*$/.test(xml.slice(offset,element.closeStart));
}

function inspectRevisionRun(run: XmlElement, textTag: "t" | "delText"):
  | { supported: true }
  | { supported: false; kind: string; reason: string } {
  let seenRPr = false;
  for (const child of run.children) {
    if (child.namespaceURI !== W_NS) {
      return { supported: false, kind: "content", reason: "revision runs must remain in the WordprocessingML namespace" };
    }
    if (child.localName === "rPr") {
      if (seenRPr) {
        return { supported: false, kind: "content", reason: "revision runs must not contain duplicate w:rPr nodes" };
      }
      seenRPr = true;
      if (child !== run.children[0]) {
        return { supported: false, kind: "content", reason: "revision run properties must remain the first child of w:r" };
      }
      continue;
    }
    if (child.localName !== textTag) {
      return {
        supported: false,
        kind: "content",
        reason: `revision runs may only contain w:rPr and w:${textTag}`,
      };
    }
  }
  return { supported: true };
}

function unsupportedRevisionContext(element: XmlElement): { supported: false; kind: string; reason: string } {
  if (hasRevisionAncestor(element)) {
    return { supported: false, kind: "nested", reason: "nested revisions are not supported" };
  }

  for (let current = element.parent; current; current = current.parent) {
    if (current.namespaceURI !== W_NS) {
      continue;
    }
    if (current.localName === "trPr" || TABLE_TAGS.has(current.localName)) {
      return { supported: false, kind: "table", reason: "table revisions are not supported" };
    }
    if (current.localName === "pPr" || current.localName === "rPr") {
      return { supported: false, kind: "format", reason: "format revisions are not supported" };
    }
  }

  return {
    supported: false,
    kind: "content",
    reason: "only run-level w:ins/w:del that are direct children of w:p are supported",
  };
}

function hasRevisionAncestor(element: XmlElement): boolean {
  for (let current = element.parent; current; current = current.parent) {
    if (current.namespaceURI === W_NS && REVISION_CONTAINER_TAGS.has(current.localName)) {
      return true;
    }
  }
  return false;
}

function revisionText(wrapper: XmlElement): string {
  const parts: string[] = [];
  for (const element of [wrapper, ...traverseDescendants(wrapper)]) {
    if (element.namespaceURI === W_NS && (element.localName === "t" || element.localName === "delText")) {
      parts.push(element.text);
    }
  }
  return parts.join("");
}

function traverseDescendants(root: XmlElement): XmlElement[] {
  const descendants: XmlElement[] = [];
  const stack = [...root.children].reverse();
  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) {
      continue;
    }
    descendants.push(current);
    for (let index = current.children.length - 1; index >= 0; index -= 1) {
      const child = current.children[index];
      if (child) {
        stack.push(child);
      }
    }
  }
  return descendants;
}

function namespaceLiftFindings(story: ScannedStoryPart, action: "accept" | "reject"): RevisionFinding[] {
  const findings: RevisionFinding[] = [];
  const keys = new Set<string>();
  for (const revision of story.revisions) {
    if (!unwrapsRevision(revision, action)) {
      continue;
    }
    const declarationConflict = namespaceLiftConflict(revision.element);
    if (declarationConflict) {
      pushFinding(findings, keys, story.part, "namespace", declarationConflict, revision.element.start);
    }
  }
  return findings;
}

function namespaceLiftConflict(wrapper: XmlElement): string | undefined {
  const declarations = namespaceDeclarations(wrapper);
  if (declarations.length === 0) {
    return undefined;
  }

  for (const child of wrapper.children) {
    for (const [name, value] of declarations) {
      const existing = child.attributes[name];
      if (existing !== undefined && existing !== value) {
        return `cannot safely lift ${name} from an unwrapped revision because child ${child.name} redeclares it`;
      }
    }
  }

  return undefined;
}

function namespaceDeclarations(element: XmlElement): Array<[string, string]> {
  return Object.entries(element.attributes).filter(
    ([name]) => element.attributeNamespaces[name] === XMLNS_NS,
  );
}

function unwrapsRevision(revision: Revision, action: "accept" | "reject"): boolean {
  return (action === "accept" && (revision.kind === "insertion" || revision.kind==='move-to'))
    || (action === "reject" && (revision.kind === "deletion" || revision.kind==='move-from'));
}

function buildStoryEdits(
  story: ScannedStoryPart,
  action: "accept" | "reject",
): Array<{ start: number; end: number; value: string }> {
  const edits: Array<{ start: number; end: number; value: string }> = [];
  for(const marker of story.moveMarkers)edits.push({start:marker.start,end:marker.end,value:''});
  for (const revision of story.revisions) {
    if(revision.kind==='move-from'||revision.kind==='move-to'){
      edits.push({start:revision.element.start,end:revision.element.end,value:unwrapsRevision(revision,action)?unwrapRevision(story.xml,revision.element,false):''});continue;
    }
    if(revision.kind==='run-properties'){
      const target=action==='accept'?revision.element:revision.element.parent!;
      edits.push({start:target.start,end:target.end,value:action==='accept'?'':restoredRunProperties(revision.element,story.xml)});
      continue;
    }
    if (action === "accept") {
      edits.push(revision.kind === "insertion"
        ? { start: revision.element.start, end: revision.element.end, value: unwrapRevision(story.xml, revision.element, false) }
        : { start: revision.element.start, end: revision.element.end, value: "" });
      continue;
    }

    edits.push(revision.kind === "insertion"
      ? { start: revision.element.start, end: revision.element.end, value: "" }
      : { start: revision.element.start, end: revision.element.end, value: unwrapRevision(story.xml, revision.element, true) });
  }
  return edits;
}

function unwrapRevision(xml: string, wrapper: XmlElement, renameDeletedText: boolean): string {
  const inner = xml.slice(wrapper.openEnd, wrapper.closeStart);
  const base = wrapper.openEnd;
  const edits: Array<{ start: number; end: number; value: string }> = [];
  const declarations = namespaceDeclarations(wrapper);

  if (declarations.length > 0) {
    for (const child of wrapper.children) {
      const missing = declarations.filter(([name]) => child.attributes[name] === undefined);
      if (missing.length === 0) {
        continue;
      }
      edits.push({
        start: child.openEnd - base - (child.selfClosing ? 2 : 1),
        end: child.openEnd - base - (child.selfClosing ? 2 : 1),
        value: missing.map(([name, value]) => ` ${name}="${escapeAttribute(value)}"`).join(""),
      });
    }
  }

  if (renameDeletedText) {
    for (const element of traverseDescendants(wrapper)) {
      if (element.namespaceURI !== W_NS || element.localName !== "delText") {
        continue;
      }
      const renamed = renameQualifiedNameLocal(element.name, "t");
      edits.push({
        start: element.start - base + 1,
        end: element.start - base + 1 + element.name.length,
        value: renamed,
      });
      if (!element.selfClosing) edits.push({
        start: element.closeStart - base + 2,
        end: element.closeStart - base + 2 + element.name.length,
        value: renamed,
      });
    }
  }

  return applyStringEdits(inner, edits);
}

function applyStringEdits(
  source: string,
  edits: Array<{ start: number; end: number; value: string }>,
): string {
  if (edits.length === 0) {
    return source;
  }

  const sorted = [...edits].sort((left, right) => left.start - right.start || left.end - right.end);
  let previousEnd = -1;
  const parts: string[] = [];
  let cursor = 0;

  for (const edit of sorted) {
    if (edit.start < 0 || edit.end < edit.start || edit.end > source.length || edit.start < previousEnd) {
      throw new OoxmlError("docx-revisions-edit-bounds", "Revision fragment edits became unsafe");
    }
    if (cursor < edit.start) {
      parts.push(source.slice(cursor, edit.start));
    }
    parts.push(edit.value);
    cursor = edit.end;
    previousEnd = edit.end;
  }

  if (cursor < source.length) {
    parts.push(source.slice(cursor));
  }
  return parts.join("");
}

function renameQualifiedNameLocal(name: string, localName: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? localName : `${name.slice(0, colon + 1)}${localName}`;
}

function describeUnsupported(findings: RevisionFinding[]): string {
  const listed = findings.map((finding) => `${finding.part}: ${finding.kind} - ${finding.reason}`);
  return `Cannot resolve DOCX revisions because the selected scope contains unsupported markup: ${listed.join("; ")}`;
}

function pushFinding(
  findings: RevisionFinding[],
  keys: Set<string>,
  part: string,
  kind: string,
  reason: string,
  offset?: number,
): void {
  const key = `${part}\u0000${kind}\u0000${reason}\u0000${offset ?? -1}`;
  if (keys.has(key)) {
    return;
  }
  keys.add(key);
  findings.push({ part, kind, reason });
}

function isProtected(pkg: OpcPackage,strict=false): boolean {
  for (const rel of pkg.relationships(pkg.mainPart())) {
    if (rel.type !== OFFICE_REL + "settings") continue;
    if(rel.external){if(strict)return true;continue;}
    const settings = parseXml(pkg.text(rel.resolved!));
    if(strict&&(settings.root.namespaceURI!==W_NS||settings.root.localName!=='settings'))return true;
    for (const protection of settings.elements) {
      if (protection.namespaceURI !== W_NS || protection.localName !== "documentProtection") continue;
      // Bounded editing refuses unknown/missing enforcement and ignores only an
      // explicit false value. Do not let the first of multiple guards hide others.
      if (!/^(?:0|false|off)$/iu.test(attribute(protection, "enforcement", W_NS) ?? "")) return true;
    }
  }
  return false;
}

function escapeAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;")
    .replaceAll("\n", "&#xA;")
    .replaceAll("\t", "&#x9;")
    .replaceAll("\r", "&#xD;");
}
