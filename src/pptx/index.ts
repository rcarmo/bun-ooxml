import { OoxmlError } from "../errors.ts";
import { OpcPackage } from "../opc/package.ts";
import { applyEdits, elements, escapeAttribute, escapeText, parseXml, type XmlElement } from "../xml/index.ts";

const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const SLIDE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

type OpenInput = string | Uint8Array | ArrayBuffer;

type StoryRun = {
  element: XmlElement;
  textElement: XmlElement;
  rPrElement?: XmlElement;
  text: string;
  start: number;
  end: number;
  attrs: Record<string, string>;
};

type StoryParagraph = {
  element: XmlElement;
  text: string;
  runs: StoryRun[];
  replaceable: boolean;
};

/**
 * Stable paragraph anchor returned by `inspectText()`.
 *
 * Invariant: callers must feed the anchor back to the exact slide instance it came from.
 * `version` pins the slide XML epoch so stale anchors refuse instead of silently re-finding.
 */
export interface TextAnchor {
  readonly kind: "pptx-text";
  readonly part: string;
  readonly paragraphIndex: number;
  readonly version: number;
  readonly text: string;
}

export type InspectedRun = {
  text: string;
  attrs: Record<string, string>;
};

export type InspectedParagraph = {
  text: string;
  runs: InspectedRun[];
  anchor: TextAnchor;
};

/**
 * Minimal PPTX read/write slice for relationship-ordered slide access and anchored text edits.
 *
 * Invariant: package bytes are the single source of truth. Slides re-parse live XML on demand
 * so transaction rollbacks in `OpcPackage` never leave stale mutable caches behind.
 */
export class Presentation {
  readonly slides: Slide[];
  readonly package: OpcPackage;

  private readonly slideVersions = new Map<string, number>();
  private readonly mainPartName: string;

  private constructor(pkg: OpcPackage) {
    this.package = pkg;
    this.mainPartName = pkg.mainPart();
    const slideParts = resolveSlideParts(pkg, this.mainPartName);
    for (const part of slideParts) {
      this.slideVersions.set(part, 0);
    }
    this.slides = slideParts.map((part, index) => new Slide(this, part, index));
  }

  static async open(input: OpenInput): Promise<Presentation> {
    const pkg = await OpcPackage.open(
      input instanceof ArrayBuffer ? new Uint8Array(input.slice(0)) : input,
    );
    return new Presentation(pkg);
  }

  async save(path: string): Promise<void> {
    await this.package.save(path);
  }

  currentSlideVersion(part: string): number {
    return this.slideVersions.get(part) ?? 0;
  }

  bumpSlideVersion(part: string): void {
    this.slideVersions.set(part, this.currentSlideVersion(part) + 1);
  }
}

/**
 * One logical slide in deck order.
 *
 * Invariant: `partName` always names an existing slide part resolved from `p:sldIdLst`; the
 * object never guesses by filename order.
 */
export class Slide {
  constructor(
    private readonly presentation: Presentation,
    readonly partName: string,
    readonly index: number,
  ) {}

  inspectText(_label: string): InspectedParagraph[] {
    const xml = this.presentation.package.text(this.partName);
    const paragraphs = collectStoryParagraphs(xml);
    const version = this.presentation.currentSlideVersion(this.partName);

    return paragraphs.map((paragraph, paragraphIndex) => ({
      text: paragraph.text,
      runs: paragraph.runs.map((run) => ({ text: run.text, attrs: { ...run.attrs } })),
      anchor: {
        kind: "pptx-text",
        part: this.partName,
        paragraphIndex,
        version,
        text: paragraph.text,
      },
    }));
  }

  /**
   * Read the existing notes part only.
   *
   * Invariant: this method never creates missing notes members. Absence is a refusal with the
   * stable machine code `PPTX_NOTES_MISSING`.
   */
  readNotesText(): string {
    const notesPart = this.presentation.package.related(this.partName, "notesSlide");
    if (!notesPart) {
      throw new OoxmlError(
        "PPTX_NOTES_MISSING",
        `Slide ${this.index + 1} has no related notes part`,
      );
    }

    return collectStoryParagraphs(this.presentation.package.text(notesPart))
      .map((paragraph) => paragraph.text)
      .filter((text) => text.length > 0)
      .join("\n");
  }

  /**
   * Replace the first exact literal match in the anchored paragraph.
   *
   * Invariant: edits are refusal-atomic. The slide part version is bumped only after the package
   * transaction commits, so stale-anchor failures leave package bytes untouched.
   */
  replaceTextAt(anchor: TextAnchor, find: string, replace: string): void {
    if (anchor.kind !== "pptx-text" || anchor.part !== this.partName) {
      throw staleAnchor(this.partName, "anchor does not belong to this slide");
    }
    if (find.length === 0) {
      throw new OoxmlError("PPTX_TEXT_EMPTY_QUERY", "Replacement text query must not be empty");
    }

    let changed = false;
    this.presentation.package.transaction(() => {
      if (anchor.version !== this.presentation.currentSlideVersion(this.partName)) {
        throw staleAnchor(this.partName, "anchor version is stale after slide mutation");
      }

      const xml = this.presentation.package.text(this.partName);
      const paragraphs = collectStoryParagraphs(xml);
      const paragraph = paragraphs[anchor.paragraphIndex];
      if (!paragraph || paragraph.text !== anchor.text) {
        throw staleAnchor(this.partName, "anchored paragraph content has changed");
      }
      if (!paragraph.replaceable) {
        throw new OoxmlError(
          "PPTX_UNSUPPORTED_TEXT_TOPOLOGY",
          `Paragraph ${anchor.paragraphIndex + 1} in ${this.partName} cannot be edited safely`,
        );
      }

      const matchStart = paragraph.text.indexOf(find);
      if (matchStart === -1) {
        throw new OoxmlError(
          "PPTX_TEXT_NOT_FOUND",
          `Text ${JSON.stringify(find)} does not occur in the anchored paragraph`,
        );
      }

      const nextXml = replaceInParagraph(xml, paragraph, matchStart, matchStart + find.length, replace);
      this.presentation.package.set(this.partName, nextXml);
      changed = true;
    });

    if (changed) {
      this.presentation.bumpSlideVersion(this.partName);
    }
  }
}

function resolveSlideParts(pkg: OpcPackage, mainPartName: string): string[] {
  const presentation = parseXml(pkg.text(mainPartName));
  if (presentation.root.localName !== "presentation" || presentation.root.namespaceURI !== PRESENTATION_NS) {
    throw new OoxmlError("PPTX_PRESENTATION_INVALID", `Invalid presentation root in ${mainPartName}`);
  }

  const sldIdList = presentation.root.children.find((child) => isElement(child, "sldIdLst", PRESENTATION_NS));
  if (!sldIdList) {
    return [];
  }

  const relationships = new Map(
    pkg.relationships(mainPartName)
      .filter((relationship) => !relationship.external)
      .map((relationship) => [relationship.id, relationship]),
  );

  return sldIdList.children
    .filter((child) => isElement(child, "sldId", PRESENTATION_NS))
    .map((slideId) => {
      const relationshipId = slideId.attributes["r:id"];
      const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
      if (!relationship || relationship.type !== SLIDE_RELATIONSHIP || !relationship.resolved) {
        throw new OoxmlError(
          "PPTX_PRESENTATION_INVALID",
          `Slide id in ${mainPartName} does not resolve to an internal slide relationship`,
        );
      }
      return relationship.resolved;
    });
}

function collectStoryParagraphs(xml: string): StoryParagraph[] {
  const document = parseXml(xml);
  const paragraphs: StoryParagraph[] = [];

  for (const body of elements(document, "txBody")) {
    for (const child of body.children) {
      if (isElement(child, "p", DRAWING_NS)) {
        paragraphs.push(analyzeParagraph(child));
      }
    }
  }

  return paragraphs;
}

function analyzeParagraph(paragraph: XmlElement): StoryParagraph {
  const runs: StoryRun[] = [];
  const textParts: string[] = [];
  let offset = 0;
  let replaceable = true;

  for (const child of paragraph.children) {
    if (isElement(child, "pPr", DRAWING_NS) || isElement(child, "endParaRPr", DRAWING_NS)) {
      continue;
    }
    if (!isElement(child, "r", DRAWING_NS)) {
      replaceable = false;
      continue;
    }

    const textElement = child.children.find((node) => isElement(node, "t", DRAWING_NS));
    const rPrElement = child.children.find((node) => isElement(node, "rPr", DRAWING_NS));
    const supportedChildren = child.children.every(
      (node) => isElement(node, "t", DRAWING_NS) || isElement(node, "rPr", DRAWING_NS),
    );

    if (!textElement || !supportedChildren || child.children.filter((node) => isElement(node, "t", DRAWING_NS)).length !== 1) {
      replaceable = false;
      continue;
    }

    const text = textElement.text;
    runs.push({
      element: child,
      textElement,
      rPrElement,
      text,
      start: offset,
      end: offset + text.length,
      attrs: { ...(rPrElement?.attributes ?? {}) },
    });
    textParts.push(text);
    offset += text.length;
  }

  return {
    element: paragraph,
    text: textParts.join(""),
    runs,
    replaceable,
  };
}

function replaceInParagraph(
  xml: string,
  paragraph: StoryParagraph,
  matchStart: number,
  matchEnd: number,
  replacement: string,
): string {
  const touched = paragraph.runs.filter((run) => run.end > matchStart && run.start < matchEnd);
  const first = touched[0];
  const last = touched.at(-1);
  if (!first || !last) {
    throw new OoxmlError("PPTX_STALE_ANCHOR", "Anchored paragraph no longer maps to text runs");
  }

  const before = first.text.slice(0, Math.max(0, matchStart - first.start));
  const after = last.text.slice(Math.max(0, matchEnd - last.start));
  const replacementRuns: string[] = [];

  if (before.length > 0) {
    replacementRuns.push(buildRunXml(xml, first, before));
  }
  if (replacement.length > 0) {
    replacementRuns.push(buildRunXml(xml, first, replacement));
  }
  if (after.length > 0) {
    replacementRuns.push(buildRunXml(xml, last, after));
  }

  return applyEdits(xml, [
    {
      start: first.element.start,
      end: last.element.end,
      value: replacementRuns.join(""),
    },
  ]);
}

function buildRunXml(xml: string, run: StoryRun, text: string): string {
  const runAttributes = serializeAttributes(run.element.attributes);
  const textAttributes = { ...run.textElement.attributes };
  if (needsPreserveSpace(text)) {
    textAttributes["xml:space"] = "preserve";
  }

  const rPrXml = run.rPrElement ? xml.slice(run.rPrElement.start, run.rPrElement.end) : "";
  return `<${run.element.name}${runAttributes}>${rPrXml}<${run.textElement.name}${serializeAttributes(textAttributes)}>${escapeText(text)}</${run.textElement.name}></${run.element.name}>`;
}

function serializeAttributes(attributes: Record<string, string>): string {
  const entries = Object.entries(attributes);
  if (entries.length === 0) {
    return "";
  }
  return entries.map(([name, value]) => ` ${name}="${escapeAttribute(value)}"`).join("");
}

function needsPreserveSpace(text: string): boolean {
  return text.length > 0 && (isXmlWhitespace(text[0]) || isXmlWhitespace(text[text.length - 1]));
}

function isXmlWhitespace(char: string | undefined): boolean {
  return char === " " || char === "\t" || char === "\n" || char === "\r";
}

function isElement(element: XmlElement, localName: string, namespaceURI?: string): boolean {
  return element.localName === localName && (namespaceURI === undefined || element.namespaceURI === namespaceURI);
}

function staleAnchor(partName: string, detail: string): OoxmlError {
  return new OoxmlError("PPTX_STALE_ANCHOR", `Stale PPTX text anchor for ${partName}: ${detail}`);
}
