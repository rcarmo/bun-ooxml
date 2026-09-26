import { OoxmlError } from "../errors.ts";
import { OpcPackage } from "../opc/index.ts";
import { applyEdits, escapeText, parseXml, type XmlDocument, type XmlElement } from "../xml/index.ts";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DOCUMENT_PART = "word/document.xml";

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

type ParagraphCollection = {
  paragraphs: ParagraphSnapshot[];
  omittedTopologies: string[];
};

/**
 * Defensive OPC package wrapper for the DOCX slice.
 *
 * Public DOCX callers keep the previous wrapper surface while the underlying
 * custody now delegates to validated, atomic `OpcPackage` behaviour.
 */
export class DocxPackage {
  constructor(private readonly opcPackage: OpcPackage) {}

  get parts(): ReadonlyMap<string, Uint8Array> {
    return new Map(this.opcPackage.names().map((name) => [name, this.opcPackage.get(name)!]));
  }

  has(name: string): boolean {
    return this.opcPackage.get(name) !== undefined;
  }

  get(name: string): Uint8Array | undefined {
    return this.opcPackage.get(name);
  }

  toBytes(): Uint8Array {
    return this.opcPackage.toBytes();
  }

  setPart(name: string, bytes: Uint8Array): void {
    this.opcPackage.set(name, bytes);
  }
}

/**
 * Snapshot handle for a document-order paragraph.
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
 * - reads `word/document.xml` paragraphs in body order, including those inside
 *   table cells and nested tables;
 * - searches exact text across direct `w:r/w:t` runs;
 * - rewrites touched text nodes in place without rebuilding unaffected runs.
 *
 * Out of scope for this slice are broader paragraph topologies such as fields,
 * revisions, content controls, text boxes and other constructs that would
 * require a more structural edit plan.
 */
export class Document {
  private version = 0;
  private xml = "";
  private xmlDocument!: XmlDocument;
  private paragraphSnapshots: ParagraphSnapshot[] = [];
  private paragraphHandles: Paragraph[] = [];
  private omittedTopologies: string[] = [];

  readonly package: DocxPackage;

  private constructor(private readonly opcPackage: OpcPackage) {
    this.package = new DocxPackage(opcPackage);
    this.reloadDocumentXml();
  }

  /** Opens a DOCX package from a filesystem path or raw bytes. */
  static async open(input: OpenInput): Promise<Document> {
    if (typeof input === "string") {
      return new Document(await OpcPackage.open(input));
    }

    const bytes = input instanceof Uint8Array
      ? cloneBytes(input)
      : new Uint8Array(input.slice(0));
    return new Document(await OpcPackage.open(bytes));
  }

  /** Current paragraph snapshot handles in document order. */
  get paragraphs(): Paragraph[] {
    return [...this.paragraphHandles];
  }

  /**
   * Finds exact matches across all searchable collected paragraphs.
   *
   * If a match is found in a searchable paragraph, unsupported paragraphs and
   * omitted blind regions are ignored for that call. If no matches are found
   * and any paragraph or omitted region is unsupported, the first refusal is
   * surfaced instead of pretending the whole document was searchable.
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

    if (matches.length > 0) {
      return matches;
    }

    const unsupported = this.paragraphHandles.find((paragraph) => !isSearchable(paragraph));
    if (unsupported) {
      throw unsupported.failure();
    }

    if (this.omittedTopologies.length > 0) {
      fail(
        "docx-unsupported-topology",
        `DOCX document search omits unsupported topology ${this.omittedTopologies[0]}`,
      );
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
    const bytes = this.opcPackage.toBytes();

    if (path) {
      await this.opcPackage.save(path);
    }

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
    this.opcPackage.set(DOCUMENT_PART, this.xml);
    this.version += 1;
    this.reloadParagraphs();
  }

  private reloadDocumentXml(): void {
    try {
      this.xml = this.opcPackage.text(DOCUMENT_PART);
    } catch (error) {
      if (error instanceof OoxmlError && error.code === "opc-part-missing") {
        fail("docx-document-part-missing", `Missing required DOCX part ${DOCUMENT_PART}`);
      }
      if (error instanceof OoxmlError && error.code === "opc-xml-encoding") {
        fail("docx-part-encoding-invalid", `${DOCUMENT_PART} has invalid XML encoding`);
      }
      throw error;
    }
    this.reloadParagraphs();
  }

  private reloadParagraphs(): void {
    this.xmlDocument = parseXml(this.xml);
    const collection = collectParagraphs(this.xmlDocument, this.version);
    this.omittedTopologies = collection.omittedTopologies;
    this.paragraphSnapshots = collection.paragraphs;
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

function collectParagraphs(document: XmlDocument, version: number): ParagraphCollection {
  const body = document.root.children.find((child) => isWord(child, "body"));
  if (!body) {
    fail("docx-document-invalid", "DOCX document XML is missing w:body");
  }

  const paragraphs: ParagraphSnapshot[] = [];
  const omittedTopologies: string[] = [];

  const collectParagraph = (paragraph: XmlElement): void => {
    paragraphs.push(analyzeParagraph(paragraph, paragraphs.length, version));
  };

  const noteOmitted = (element: XmlElement): void => {
    omittedTopologies.push(element.name);
  };

  const collectTableCell = (cell: XmlElement): void => {
    for (const child of cell.children) {
      if (isWord(child, "tcPr")) {
        continue;
      }
      if (isWord(child, "p")) {
        collectParagraph(child);
        continue;
      }
      if (isWord(child, "tbl")) {
        collectTable(child);
        continue;
      }
      noteOmitted(child);
    }
  };

  const collectTableRow = (row: XmlElement): void => {
    for (const child of row.children) {
      if (isWord(child, "trPr")) {
        continue;
      }
      if (isWord(child, "tc")) {
        collectTableCell(child);
        continue;
      }
      noteOmitted(child);
    }
  };

  const collectTable = (table: XmlElement): void => {
    for (const child of table.children) {
      if (isWord(child, "tblPr") || isWord(child, "tblGrid") || isWord(child, "tblPrEx")) {
        continue;
      }
      if (isWord(child, "tr")) {
        collectTableRow(child);
        continue;
      }
      noteOmitted(child);
    }
  };

  for (const child of body.children) {
    if (isWord(child, "p")) {
      collectParagraph(child);
      continue;
    }
    if (isWord(child, "tbl")) {
      collectTable(child);
      continue;
    }
    if (isWord(child, "sectPr")) {
      continue;
    }
    noteOmitted(child);
  }

  return { paragraphs, omittedTopologies };
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
