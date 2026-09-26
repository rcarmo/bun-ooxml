import { OoxmlError } from "../errors.ts";
import { readZip, writeZip } from "../opc/zip.ts";
import { applyEdits, escapeText, parseXml, type XmlDocument, type XmlElement } from "../xml/index.ts";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DOCUMENT_PART = "word/document.xml";
const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });
const UTF8_ENCODER = new TextEncoder();

type OpenInput = string | Uint8Array | ArrayBuffer;

type TextSegment = {
  element: XmlElement;
  text: string;
  paragraphStart: number;
  paragraphEnd: number;
};

type ParagraphSnapshot = {
  index: number;
  version: number;
  text: string;
  segments: TextSegment[];
  searchable: boolean;
  unsupported?: string;
};

type SpanSnapshot = {
  version: number;
  paragraphIndex: number;
  start: number;
  end: number;
  query: string;
};

/**
 * Defensive OPC package wrapper for the DOCX slice.
 *
 * All byte arrays are cloned on ingress and egress so callers cannot mutate the
 * in-memory package behind the document's back.
 */
export class DocxPackage {
  private readonly partsMap: Map<string, Uint8Array>;

  constructor(parts: ReadonlyMap<string, Uint8Array>) {
    this.partsMap = new Map(
      Array.from(parts, ([name, bytes]) => [name, cloneBytes(bytes)]),
    );
  }

  get parts(): ReadonlyMap<string, Uint8Array> {
    return new Map(
      Array.from(this.partsMap, ([name, bytes]) => [name, cloneBytes(bytes)]),
    );
  }

  has(name: string): boolean {
    return this.partsMap.has(name);
  }

  get(name: string): Uint8Array | undefined {
    const bytes = this.partsMap.get(name);
    return bytes ? cloneBytes(bytes) : undefined;
  }

  toBytes(): Uint8Array {
    return writeZip(this.partsMap);
  }

  setPart(name: string, bytes: Uint8Array): void {
    this.partsMap.set(name, cloneBytes(bytes));
  }
}

/**
 * Snapshot handle for a body paragraph.
 *
 * This first slice only treats paragraphs made of direct `w:r/w:t` text runs as
 * searchable. Any broader topology is refused instead of flattened lossy.
 * Paragraph handles become stale after any successful document mutation.
 */
export class Paragraph {
  constructor(
    private readonly documentRef: Document,
    private readonly snapshot: ParagraphSnapshot,
  ) {}

  get index(): number {
    return this.snapshot.index;
  }

  /**
   * Flattened paragraph text for exact matching.
   *
   * Refuses stale handles and unsupported paragraph topologies.
   */
  get text(): string {
    this.ensureFresh();
    this.ensureSearchable();
    return this.snapshot.text;
  }

  /**
   * Returns exact, non-overlapping matches against the flattened paragraph text.
   *
   * Assumptions/invariants:
   * - the query must be non-empty;
   * - the paragraph handle must be fresh for the current document version;
   * - unsupported paragraph topologies are refused with a stable error code.
   */
  find(query: string): Span[] {
    this.ensureFresh();
    this.ensureSearchable();
    if (query.length === 0) {
      fail("docx-empty-query", "DOCX exact search text must not be empty");
    }

    const matches: Span[] = [];
    let cursor = 0;
    while (cursor <= this.snapshot.text.length - query.length) {
      const start = this.snapshot.text.indexOf(query, cursor);
      if (start === -1) {
        break;
      }
      matches.push(
        new Span(this.documentRef, {
          version: this.snapshot.version,
          paragraphIndex: this.snapshot.index,
          start,
          end: start + query.length,
          query,
        }),
      );
      cursor = start + Math.max(query.length, 1);
    }

    return matches;
  }

  /**
   * Materialises the refusal used when this slice encounters a topology it does
   * not yet know how to search or edit safely.
   */
  failure(): OoxmlError {
    return new OoxmlError(
      "docx-unsupported-topology",
      `Paragraph ${this.snapshot.index + 1} uses unsupported DOCX topology ${this.snapshot.unsupported ?? "later-slice-content"}`,
    );
  }

  private ensureFresh(): void {
    if (this.documentRef.currentVersion() !== this.snapshot.version) {
      fail("docx-stale-paragraph", "DOCX paragraph handle is stale after document mutation");
    }
  }

  private ensureSearchable(): void {
    if (!this.snapshot.searchable) {
      throw this.failure();
    }
  }
}

/**
 * Exact-match handle captured from a paragraph search.
 *
 * A span is tied to a specific document version and original text range.
 * `text` is the remembered query payload; mutation is revalidated at replace
 * time so stale spans refuse safely before changing package bytes.
 */
export class Span {
  constructor(
    private readonly documentRef: Document,
    private readonly snapshot: SpanSnapshot,
  ) {}

  get text(): string {
    return this.snapshot.query;
  }

  /**
   * Replaces the original exact-match range.
   *
   * Refusal invariants:
   * - stale spans are rejected with `docx-stale-span`;
   * - unsupported paragraph topologies are rejected with
   *   `docx-unsupported-topology`;
   * - on refusal, document XML and package bytes remain unchanged.
   */
  async replace(replacement: string): Promise<void> {
    this.documentRef.replaceSpan(this.snapshot, replacement);
  }
}

/**
 * First DOCX text-editing slice.
 *
 * Scope for this slice:
 * - reads only `word/document.xml` body paragraphs;
 * - searches exact text across direct `w:r/w:t` runs;
 * - rewrites touched text nodes in place without rebuilding unaffected runs.
 *
 * Out of scope for this slice are broader paragraph topologies such as fields,
 * revisions, content controls and other constructs that would require a more
 * structural edit plan.
 */
export class Document {
  private originalBytes?: Uint8Array;
  private dirty = false;
  private version = 0;
  private xml = "";
  private xmlDocument!: XmlDocument;
  private paragraphSnapshots: ParagraphSnapshot[] = [];
  private paragraphHandles: Paragraph[] = [];

  readonly package: DocxPackage;

  private constructor(parts: ReadonlyMap<string, Uint8Array>, originalBytes?: Uint8Array) {
    this.package = new DocxPackage(parts);
    this.originalBytes = originalBytes ? cloneBytes(originalBytes) : undefined;
    this.reloadDocumentXml();
  }

  /** Opens a DOCX package from a filesystem path or raw bytes. */
  static async open(input: OpenInput): Promise<Document> {
    if (typeof input === "string") {
      const bytes = new Uint8Array(await Bun.file(input).arrayBuffer());
      return new Document(readZip(bytes), bytes);
    }

    const bytes = input instanceof Uint8Array
      ? cloneBytes(input)
      : new Uint8Array(input.slice(0));
    return new Document(readZip(bytes), bytes);
  }

  /** Current paragraph snapshot handles in document order. */
  get paragraphs(): Paragraph[] {
    return [...this.paragraphHandles];
  }

  /**
   * Finds exact matches across all searchable body paragraphs.
   *
   * If a match is found in a searchable paragraph, unsupported paragraphs are
   * ignored for that call. If no matches are found and any paragraph is
   * unsupported, the first refusal is surfaced instead of pretending the whole
   * document was searchable.
   */
  find(query: string): Span[] {
    const matches = this.paragraphHandles.flatMap((paragraph) => {
      try {
        return paragraph.find(query);
      } catch (error) {
        if (error instanceof OoxmlError && error.code === "docx-unsupported-topology") {
          return [];
        }
        throw error;
      }
    });

    if (matches.length > 0 || this.paragraphHandles.every((paragraph) => isSearchable(paragraph))) {
      return matches;
    }

    const unsupported = this.paragraphHandles.find((paragraph) => !isSearchable(paragraph));
    if (unsupported) {
      throw unsupported.failure();
    }
    return matches;
  }

  /**
   * Serialises the current package.
   *
   * Saving never writes back to the original source path implicitly; callers
   * must pass `path` when they want filesystem output. This keeps test fixtures
   * and other source corpora safe across runs.
   */
  async save(path?: string): Promise<Uint8Array> {
    const bytes = this.dirty
      ? this.package.toBytes()
      : this.originalBytes
        ? cloneBytes(this.originalBytes)
        : this.package.toBytes();

    if (path) {
      await Bun.write(path, bytes);
    }

    this.originalBytes = cloneBytes(bytes);
    this.dirty = false;
    return cloneBytes(bytes);
  }

  currentVersion(): number {
    return this.version;
  }

  /**
   * Internal exact-range replacement primitive used by `Span.replace()`.
   *
   * Mutation invariants:
   * - the span version must match the live document version;
   * - the paragraph must still be present and searchable;
   * - the current flattened text at the captured range must still equal the
   *   original query;
   * - every touched text node is edited in place and `xml:space="preserve"`
   *   is added when the replacement would otherwise lose boundary whitespace.
   *
   * Any refusal aborts before XML/package mutation.
   */
  replaceSpan(span: SpanSnapshot, replacement: string): void {
    if (this.version !== span.version) {
      fail("docx-stale-span", "DOCX span is stale after document mutation");
    }

    const paragraph = this.paragraphSnapshots[span.paragraphIndex];
    if (!paragraph) {
      fail("docx-paragraph-missing", `Missing paragraph ${span.paragraphIndex + 1}`);
    }
    if (!paragraph.searchable) {
      throw new Paragraph(this, paragraph).failure();
    }
    if (paragraph.text.slice(span.start, span.end) !== span.query) {
      fail("docx-stale-span", "DOCX span text no longer matches the original target");
    }

    const touched = paragraph.segments.filter(
      (segment) => segment.paragraphEnd > span.start && segment.paragraphStart < span.end,
    );
    if (touched.length === 0) {
      fail("docx-stale-span", "DOCX span no longer maps to paragraph text segments");
    }

    const edits: Array<{ start: number; end: number; value: string }> = [];
    let remaining = replacement;

    for (const [index, segment] of touched.entries()) {
      const localStart = Math.max(0, span.start - segment.paragraphStart);
      const localEnd = Math.min(segment.text.length, span.end - segment.paragraphStart);
      const covered = localEnd - localStart;
      const chunk = index === touched.length - 1
        ? remaining
        : remaining.slice(0, Math.min(covered, remaining.length));
      remaining = index === touched.length - 1
        ? ""
        : remaining.slice(Math.min(covered, remaining.length));

      const nextText =
        segment.text.slice(0, localStart) + chunk + segment.text.slice(localEnd);

      if (needsPreserveSpace(nextText) && segment.element.attributes["xml:space"] !== "preserve") {
        edits.push({
          start: segment.element.openEnd - 1,
          end: segment.element.openEnd - 1,
          value: ' xml:space="preserve"',
        });
      }
      edits.push({
        start: segment.element.openEnd,
        end: segment.element.closeStart,
        value: escapeText(nextText),
      });
    }

    this.xml = applyEdits(this.xml, edits);
    this.package.setPart(DOCUMENT_PART, UTF8_ENCODER.encode(this.xml));
    this.dirty = true;
    this.version += 1;
    this.reloadParagraphs();
  }

  private reloadDocumentXml(): void {
    const bytes = this.package.get(DOCUMENT_PART);
    if (!bytes) {
      fail("docx-document-part-missing", `Missing required DOCX part ${DOCUMENT_PART}`);
    }
    try {
      this.xml = UTF8_DECODER.decode(bytes);
    } catch {
      fail("docx-part-encoding-invalid", `${DOCUMENT_PART} is not valid UTF-8`);
    }
    this.reloadParagraphs();
  }

  private reloadParagraphs(): void {
    this.xmlDocument = parseXml(this.xml);
    this.paragraphSnapshots = collectParagraphs(this.xmlDocument, this.version);
    this.paragraphHandles = this.paragraphSnapshots.map(
      (snapshot) => new Paragraph(this, snapshot),
    );
  }
}

export async function open(input: OpenInput): Promise<Document> {
  return Document.open(input);
}

export async function save(document: Document, path?: string): Promise<Uint8Array> {
  return document.save(path);
}

function collectParagraphs(document: XmlDocument, version: number): ParagraphSnapshot[] {
  const body = document.root.children.find((child) => isWord(child, "body"));
  if (!body) {
    fail("docx-document-invalid", "DOCX document XML is missing w:body");
  }

  return body.children
    .filter((child) => isWord(child, "p"))
    .map((paragraph, index) => analyzeParagraph(paragraph, index, version));
}

function analyzeParagraph(paragraph: XmlElement, index: number, version: number): ParagraphSnapshot {
  const segments: TextSegment[] = [];
  const textParts: string[] = [];
  let paragraphOffset = 0;
  let unsupported: string | undefined;

  for (const child of paragraph.children) {
    if (isWord(child, "pPr")) {
      continue;
    }
    if (!isWord(child, "r")) {
      unsupported = unsupported ?? child.name;
      continue;
    }

    for (const runChild of child.children) {
      if (isWord(runChild, "rPr")) {
        continue;
      }
      if (!isWord(runChild, "t") || runChild.children.length > 0) {
        unsupported = unsupported ?? runChild.name;
        continue;
      }

      const text = runChild.text;
      segments.push({
        element: runChild,
        text,
        paragraphStart: paragraphOffset,
        paragraphEnd: paragraphOffset + text.length,
      });
      textParts.push(text);
      paragraphOffset += text.length;
    }
  }

  return {
    index,
    version,
    text: textParts.join(""),
    segments,
    searchable: unsupported === undefined,
    unsupported,
  };
}

function isSearchable(paragraph: Paragraph): boolean {
  try {
    void paragraph.text;
    return true;
  } catch (error) {
    return !(error instanceof OoxmlError) || error.code !== "docx-unsupported-topology";
  }
}

function isWord(element: XmlElement, localName: string): boolean {
  return element.localName === localName && element.namespaceURI === W_NS;
}

function needsPreserveSpace(text: string): boolean {
  return text.length > 0 && (isXmlWhitespace(text[0]) || isXmlWhitespace(text[text.length - 1]));
}

function isXmlWhitespace(char: string | undefined): boolean {
  return char === " " || char === "\t" || char === "\n" || char === "\r";
}

function cloneBytes(bytes: Uint8Array): Uint8Array {
  return Uint8Array.from(bytes);
}

function fail(code: string, message: string): never {
  throw new OoxmlError(code, message);
}
