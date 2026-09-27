import { OoxmlError } from "../errors.ts";
import {appendParagraphRun} from './append-run.ts';
import {readRowHeader,editRowHeader} from './row-header.ts';
import {readCellProperties,editCellProperties,normalizeCellProperties,type DirectCellProperties,type CellPropertiesPatch} from './cell-properties.ts';
export type {DirectCellProperties,CellPropertiesPatch,CellTopBorder} from './cell-properties.ts';
import {inspectEffectiveFormatting,type EffectiveRunFormatting} from './effective-formatting.ts';
export type {EffectiveRunFormatting,EffectiveFlag,FormattingContribution} from './effective-formatting.ts';
import {formatRunProperties,directFontSizes as readDirectFontSizes,directFontNames as readDirectFontNames,directRunFlags as readDirectRunFlags,directRunAppearance as readDirectRunAppearance,type DirectRunFlags,type DirectRunAppearance,type DirectRunPatch} from './run-formatting.ts';
export type {DirectRunFlags,DirectRunAppearance,UnderlineStyle,HighlightColor,RunVerticalAlignment} from './run-formatting.ts';
import {readPageLayout,replacePageLayout,normalizePageLayout,type PageLayout} from './page-layout.ts';
export type {PageLayout} from './page-layout.ts';
import {directParagraphStyle,replaceParagraphStyle} from './paragraph-style.ts';
import {readParagraphProperties,editParagraphProperties,normalizeParagraphProperties,type DirectParagraphProperties,type ParagraphPropertiesPatch} from './paragraph-properties.ts';
export type {DirectParagraphProperties,ParagraphPropertiesPatch,ParagraphAlignment} from './paragraph-properties.ts';
import {authorParagraphStyle,normalizeStyleOptions,type AddParagraphStyleOptions,type ParagraphStyleDefinitionReceipt} from './style-authoring.ts';
export type {AddParagraphStyleOptions,ParagraphStyleDefinitionReceipt} from './style-authoring.ts';
import { OpcPackage, getContentType } from "../opc/index.ts";
import { applyEdits, attribute, elements, escapeAttribute, escapeText, parseXml, type XmlDocument, type XmlElement } from "../xml/index.ts";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_PART = "word/document.xml";
const DOCUMENT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const UTF8_ENCODER = new TextEncoder();
const MAX_TABLE_ROWS = 100;
const MAX_TABLE_COLUMNS = 100;
const MAX_TABLE_CELLS = 10_000;
const DEFAULT_TABLE_WIDTH_DXA = 8_640;

type OpenInput = string | Uint8Array | ArrayBuffer;

export type AddParagraphOptions = {
  bold?: boolean;
  italic?: boolean;
  style?: string;
};

/** Direct overrides only: null removes the property, false is explicit off. */
export type RunFormattingPatch = DirectRunPatch;
export type RunFormattingReceipt = { changedRuns: number };

type NormalizedAddParagraphOptions = {
  bold: boolean;
  italic: boolean;
  style?: string;
};

type TextSegment = {
  element: XmlElement;
  text: string;
  paragraphStart: number;
  paragraphEnd: number;
};

type ParagraphSnapshot = {
  element: XmlElement;
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

type TableParagraphSnapshot = {
  text: string;
  searchable: boolean;
  unsupported?: string;
};

type TableCellSnapshot = {
  version: number;
  tableIndex: number;
  row: number;
  column: number;
  element: XmlElement;
  paragraphs: TableParagraphSnapshot[];
  unsupported?: string;
};

type TableGridEntry =
  | { kind: "cell"; cell: TableCellSnapshot }
  | { kind: "merged" };

type TableSnapshot = {
  index: number;
  version: number;
  rows: number;
  columns: number;
  element: XmlElement;
  cells: Array<Array<TableGridEntry | undefined>>;
  unsupported?: string;
};

type TableCollection = {
  tables: TableSnapshot[];
};

type DocumentCollections = {
  paragraphs: ParagraphCollection;
  tables: TableCollection;
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

  /** Cumulative payload changes relative to the opened/created baseline. */
  diff(): import("../opc/package.ts").PackageDiff { return this.opcPackage.diff(); }

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

  /** Bounded bold/italic cascade with provenance; unsupported contexts refuse. */
  effectiveRunFormatting(): EffectiveRunFormatting[] {
    this.ensureFresh();
    return this.documentRef.inspectParagraphFormatting(this.snapshot);
  }

  /** Append one independent plain-text run and return a fresh paragraph handle. */
  appendRun(text: string, formatting: RunFormattingPatch = {}): Paragraph {
    this.ensureFresh(); this.ensureSearchable();
    return this.documentRef.appendRunToParagraph(this.snapshot,text,formatting);
  }

  /** Direct paragraph values only; null means no direct property. */
  directProperties(): DirectParagraphProperties {
    this.ensureFresh(); return this.documentRef.inspectParagraphProperties(this.snapshot);
  }

  setProperties(patch: ParagraphPropertiesPatch): {changed:number} {
    this.ensureFresh(); return this.documentRef.setParagraphProperties(this.snapshot, patch);
  }

  /** Direct paragraph style ID, without evaluating inheritance. */
  get styleId(): string | undefined { this.ensureFresh(); return this.documentRef.paragraphStyle(this.snapshot); }

  setStyle(styleId: string | null): {changed:number} {
    this.ensureFresh();return this.documentRef.setParagraphStyle(this.snapshot,styleId);
  }

  /** Direct non-complex-script sizes in run order; null means no direct size. */
  directFontSizes(): Array<number|null> {
    this.ensureFresh();return this.documentRef.inspectParagraphFontSizes(this.snapshot);
  }

  /** Matching direct ASCII/high-ANSI font names; theme/script metadata refuses. */
  directFontNames(): Array<string|null> {
    this.ensureFresh(); return this.documentRef.inspectParagraphFontNames(this.snapshot);
  }

  /** Read direct Boolean overrides in run order; null means absent. */
  directRunFlags(): DirectRunFlags[] {
    this.ensureFresh(); return this.documentRef.inspectParagraphRunFlags(this.snapshot);
  }

  /** Direct scalar overrides in run order; theme-dependent values refuse. */
  directRunAppearance(): DirectRunAppearance[] {
    this.ensureFresh(); return this.documentRef.inspectParagraphRunAppearance(this.snapshot);
  }

  /** Apply direct formatting overrides to every supported direct run. */
  setRunFormatting(patch: RunFormattingPatch): RunFormattingReceipt {
    this.ensureFresh();
    return this.documentRef.formatParagraphRuns(this.snapshot, patch);
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

/** Snapshot handle for a top-level DOCX table. */
export class Table {
  constructor(
    private readonly documentRef: Document,
    private readonly snapshot: TableSnapshot,
  ) {}

  get rows(): number {
    const table = this.liveTable();
    return table.rows;
  }

  get columns(): number {
    const table = this.liveTable();
    return table.columns;
  }

  /** Inspect the direct row marker, not effective pagination or style inheritance. */
  isRowHeader(row: number): boolean {
    return this.documentRef.getRowHeader(this.liveTable(),row);
  }

  /** Boolean writes an explicit marker; null removes it. Table handles remain usable. */
  setRowHeader(row: number, value: boolean | null): {changed: number} {
    return this.documentRef.setRowHeader(this.liveTable(),row,value);
  }

  cell(row: number, column: number): TableCell {
    const table = this.liveTable();
    assertTableCoordinate(table, row, column);
    if (table.unsupported) {
      fail("docx-table-unsupported", `DOCX table ${table.index + 1} uses unsupported topology ${table.unsupported}`);
    }

    const entry = table.cells[row]?.[column];
    if (!entry) {
      fail("docx-table-unsupported", `DOCX table ${this.snapshot.index + 1} cannot resolve cell (${row},${column}) safely`);
    }
    if (entry.kind === "merged") {
      fail("docx-table-merged-cell", `DOCX table ${this.snapshot.index + 1} cell (${row},${column}) is merged`);
    }

    return new TableCell(this.documentRef, entry.cell);
  }

  private liveTable(): TableSnapshot {
    this.ensureFresh();
    return this.documentRef.resolveTable(this.snapshot);
  }

  private ensureFresh(): void {
    if (this.documentRef.currentTableVersion() !== this.snapshot.version) {
      fail("docx-stale-table", "DOCX table handle is stale after document mutation");
    }
  }
}

/** Snapshot handle for a simple editable table cell. */
export class TableCell {
  constructor(
    private readonly documentRef: Document,
    private readonly snapshot: TableCellSnapshot,
  ) {}

  get text(): string {
    this.ensureFresh();
    this.ensureSupported();
    return this.snapshot.paragraphs.map((paragraph) => paragraph.text).join("\n");
  }

  directProperties(): DirectCellProperties {
    this.ensureFresh(); this.ensureSupported(); return this.documentRef.inspectCellProperties(this.snapshot);
  }

  setProperties(patch: CellPropertiesPatch): {changed:number} {
    this.ensureFresh(); this.ensureSupported(); return this.documentRef.setCellProperties(this.snapshot,patch);
  }

  set text(value: string) {
    this.ensureFresh();
    this.ensureSupported();
    if (typeof value !== "string") {
      fail("docx-invalid-argument", "DOCX table cell text must be a string");
    }
    this.documentRef.replaceTableCellText(this.snapshot, value);
  }

  private ensureFresh(): void {
    if (this.documentRef.currentVersion() !== this.snapshot.version) {
      fail("docx-stale-table-cell", "DOCX table cell handle is stale after document mutation");
    }
  }

  private ensureSupported(): void {
    if (this.snapshot.unsupported) {
      fail(
        "docx-table-cell-unsupported",
        `DOCX table cell (${this.snapshot.row},${this.snapshot.column}) uses unsupported topology ${this.snapshot.unsupported}`,
      );
    }
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
  private tableVersion = 0;
  private xml = "";
  private xmlDocument!: XmlDocument;
  private paragraphSnapshots: ParagraphSnapshot[] = [];
  private paragraphHandles: Paragraph[] = [];
  private omittedTopologies: string[] = [];
  private tableSnapshots: TableSnapshot[] = [];
  private tableHandles: Table[] = [];

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

  /** Creates a new minimal DOCX package without borrowed template bytes. */
  static create(): Document {
    return new Document(OpcPackage.fromParts(buildMinimalDocxParts()));
  }

  /** Author a named paragraph style without computing inherited formatting. */
  addParagraphStyle(styleId: string, options: AddParagraphStyleOptions): ParagraphStyleDefinitionReceipt {
    const normalized=normalizeStyleOptions(styleId,options);
    if(this.opcPackage.text(DOCUMENT_PART)!==this.xml)fail('docx-stale-document','Document XML changed outside this document handle');
    this.assertFormattingUnprotected();
    // Style definitions do not alter paragraph offsets; existing handles remain usable.
    return authorParagraphStyle(this.opcPackage,DOCUMENT_PART,styleId,normalized);
  }

  /** Direct page geometry of the existing final body section. */
  getPageLayout(): PageLayout {
    if(this.opcPackage.text(DOCUMENT_PART)!==this.xml)fail('docx-stale-document','Document changed outside its wrapper');
    return readPageLayout(this.xml);
  }
  setPageLayout(layout:PageLayout):{changed:number} {
    const normalized=normalizePageLayout(layout);
    if(this.opcPackage.text(DOCUMENT_PART)!==this.xml)fail('docx-stale-document','Document changed outside its wrapper');
    this.assertFormattingUnprotected();
    for(const rel of this.opcPackage.relationships(DOCUMENT_PART))if(rel.type==='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'&&!rel.external){const settings=parseXml(this.opcPackage.text(rel.resolved!));if(['mirrorMargins','gutterAtTop','bookFoldPrinting','bookFoldRevPrinting'].some(name=>elements(settings,name,W_NS).length))fail('docx-layout-unsupported','Mirrored/book-fold/gutter settings require computed page geometry');}
    const next=replacePageLayout(this.xml,normalized);if(next===this.xml)return {changed:0};
    this.commitParagraphProperties(next);return {changed:1};
  }

  /** Current paragraph snapshot handles in document order. */
  get paragraphs(): Paragraph[] {
    return [...this.paragraphHandles];
  }

  /** Current top-level table snapshot handles in document order. */
  get tables(): Table[] {
    return [...this.tableHandles];
  }

  /**
   * Appends a simple paragraph immediately before the body section properties.
   *
   * Supported creation surface for this slice:
   * - plain text only, with XML-escaped text nodes;
   * - optional bold/italic run properties on the single authored run;
   * - optional paragraph style id when the existing styles part resolves and
   *   contains that paragraph style.
   *
   * Any refusal happens before document or package mutation.
   */
  addParagraph(text: string = "", options?: AddParagraphOptions): Paragraph {
    const normalizedOptions = normalizeAddParagraphOptions(options);
    if (typeof text !== "string") {
      fail("docx-invalid-argument", "DOCX paragraph text must be a string");
    }
    if (normalizedOptions.style) {
      this.assertSupportedParagraphStyle(normalizedOptions.style);
    }

    const body = this.requireBody();
    const sectionProperties = body.children.find((child) => isWord(child, "sectPr"));
    const insertAt = sectionProperties?.start ?? body.closeStart;
    const paragraphXml = buildParagraphXml(text, normalizedOptions);
    const nextXml = applyEdits(this.xml, [{ start: insertAt, end: insertAt, value: paragraphXml }]);
    const nextVersion = this.version + 1;
    const nextTableVersion = this.tableVersion + 1;
    const nextDocument = parseXml(nextXml);
    const nextCollections = collectDocumentCollections(nextDocument, nextVersion, nextTableVersion);
    const nextIndex = this.paragraphSnapshots.length;

    this.opcPackage.transaction(() => {
      this.opcPackage.set(DOCUMENT_PART, nextXml);
    });

    this.xml = nextXml;
    this.version = nextVersion;
    this.tableVersion = nextTableVersion;
    this.xmlDocument = nextDocument;
    this.applyDocumentCollections(nextCollections);

    const paragraph = this.paragraphHandles[nextIndex];
    if (!paragraph) {
      fail("docx-document-invalid", "DOCX paragraph append did not materialise the authored paragraph");
    }
    return paragraph;
  }

  /**
   * Appends a simple rectangular top-level table before the body section properties.
   *
   * Authored tables are deliberately bounded and use a plain rectangular grid.
   * Any refusal happens before document or package mutation.
   */
  addTable(rows: number, columns: number): Table {
    assertTableDimensions(rows, columns);

    const body = this.requireBody();
    const sectionProperties = body.children.find((child) => isWord(child, "sectPr"));
    const insertAt = sectionProperties?.start ?? body.closeStart;
    const tableXml = buildTableXml(rows, columns);
    const nextXml = applyEdits(this.xml, [{ start: insertAt, end: insertAt, value: tableXml }]);
    const nextVersion = this.version + 1;
    const nextTableVersion = this.tableVersion + 1;
    const nextDocument = parseXml(nextXml);
    const nextCollections = collectDocumentCollections(nextDocument, nextVersion, nextTableVersion);
    const nextIndex = this.tableSnapshots.length;

    this.opcPackage.transaction(() => {
      this.opcPackage.set(DOCUMENT_PART, nextXml);
    });

    this.xml = nextXml;
    this.version = nextVersion;
    this.tableVersion = nextTableVersion;
    this.xmlDocument = nextDocument;
    this.applyDocumentCollections(nextCollections);

    const table = this.tableHandles[nextIndex];
    if (!table) {
      fail("docx-document-invalid", "DOCX table append did not materialise the authored table");
    }
    return table;
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

  currentTableVersion(): number {
    return this.tableVersion;
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

    const nextXml = applyEdits(this.xml, edits);
    const nextVersion = this.version + 1;
    const nextDocument = parseXml(nextXml);
    const nextCollections = collectDocumentCollections(nextDocument, nextVersion, this.tableVersion);

    this.opcPackage.transaction(() => {
      this.opcPackage.set(DOCUMENT_PART, nextXml);
    });

    this.xml = nextXml;
    this.version = nextVersion;
    this.xmlDocument = nextDocument;
    this.applyDocumentCollections(nextCollections);
  }

  inspectParagraphFormatting(snapshot: ParagraphSnapshot): EffectiveRunFormatting[] {
    this.assertParagraphSnapshot(snapshot);
    return inspectEffectiveFormatting(this.opcPackage,DOCUMENT_PART,snapshot.element);
  }

  inspectParagraphFontNames(snapshot: ParagraphSnapshot): Array<string|null> {
    this.assertParagraphSnapshot(snapshot);
    if(!snapshot.searchable)throw new Paragraph(this,snapshot).failure();
    return readDirectFontNames(this.xml,snapshot.element);
  }

  inspectParagraphRunAppearance(snapshot: ParagraphSnapshot): DirectRunAppearance[] {
    this.assertParagraphSnapshot(snapshot);
    if(!snapshot.searchable)throw new Paragraph(this,snapshot).failure();
    return readDirectRunAppearance(this.xml,snapshot.element);
  }

  inspectParagraphRunFlags(snapshot: ParagraphSnapshot): DirectRunFlags[] {
    this.assertParagraphSnapshot(snapshot);
    if(!snapshot.searchable)throw new Paragraph(this,snapshot).failure();
    return readDirectRunFlags(this.xml,snapshot.element);
  }

  inspectParagraphFontSizes(snapshot: ParagraphSnapshot): Array<number|null> {
    this.assertParagraphSnapshot(snapshot);
    if(!snapshot.searchable)throw new Paragraph(this,snapshot).failure();
    return readDirectFontSizes(this.xml,snapshot.element);
  }

  inspectParagraphProperties(snapshot: ParagraphSnapshot): DirectParagraphProperties {
    this.assertParagraphSnapshot(snapshot); return readParagraphProperties(this.xml,snapshot.element);
  }

  setParagraphProperties(snapshot: ParagraphSnapshot, input: ParagraphPropertiesPatch): {changed:number} {
    this.assertParagraphSnapshot(snapshot);
    const patch=normalizeParagraphProperties(input);
    if(!snapshot.searchable)throw new Paragraph(this,snapshot).failure();
    this.assertFormattingUnprotected();
    const next=editParagraphProperties(this.xml,snapshot.element,patch);
    if(next===this.xml)return {changed:0};
    this.commitParagraphProperties(next);return {changed:1};
  }

  paragraphStyle(snapshot: ParagraphSnapshot): string | undefined {
    this.assertParagraphSnapshot(snapshot);
    return directParagraphStyle(this.xml,snapshot.element);
  }

  setParagraphStyle(snapshot: ParagraphSnapshot, styleId: string | null): {changed:number} {
    this.assertParagraphSnapshot(snapshot);
    if(styleId!==null&&(typeof styleId!=='string'||!styleId.trim()))fail('docx-style-argument','Expected a nonempty style ID or null');
    if(!snapshot.searchable)throw new Paragraph(this,snapshot).failure();
    this.assertFormattingUnprotected();
    if(styleId!==null)this.assertSupportedParagraphStyle(styleId);
    const next=replaceParagraphStyle(this.xml,snapshot.element,styleId);
    if(next===this.xml)return {changed:0};
    this.commitParagraphProperties(next);
    return {changed:1};
  }

  formatParagraphRuns(snapshot: ParagraphSnapshot, patch: RunFormattingPatch): RunFormattingReceipt {
    this.assertParagraphSnapshot(snapshot);
    if(!snapshot.searchable)throw new Paragraph(this,snapshot).failure();
    this.assertFormattingUnprotected();
    const result=formatRunProperties(this.xml,snapshot.element,patch);
    if(!result.changedRuns)return {changedRuns:0};
    this.commitParagraphProperties(result.xml);
    return {changedRuns:result.changedRuns};
  }

  appendRunToParagraph(snapshot:ParagraphSnapshot,text:string,formatting:RunFormattingPatch):Paragraph {
    this.assertParagraphSnapshot(snapshot);
    if(!snapshot.searchable)throw new Paragraph(this,snapshot).failure();
    this.assertFormattingUnprotected();
    const nextXml=appendParagraphRun(this.xml,snapshot.element,text,formatting);
    const nextVersion=this.version+1,nextDocument=parseXml(nextXml);
    const collections=collectDocumentCollections(nextDocument,nextVersion,this.tableVersion);
    const expected=this.paragraphSnapshots.map((p,i)=>i===snapshot.index?p.text+text:p.text);
    if(JSON.stringify(collections.paragraphs.paragraphs.map(p=>p.text))!==JSON.stringify(expected))fail('docx-run-unsafe','Appending a run changed unexpected paragraph text');
    const grids=(tables:TableSnapshot[])=>tables.map(t=>[t.rows,t.columns]);
    if(JSON.stringify(grids(collections.tables.tables))!==JSON.stringify(grids(this.tableSnapshots)))fail('docx-run-unsafe','Appending a run changed the table grid');
    this.opcPackage.transaction(()=>{this.opcPackage.set(DOCUMENT_PART,nextXml);this.opcPackage.toBytes();});
    this.xml=nextXml;this.version=nextVersion;this.xmlDocument=nextDocument;this.applyDocumentCollections(collections);
    return this.paragraphHandles[snapshot.index]!;
  }

  private assertParagraphSnapshot(snapshot: ParagraphSnapshot): void {
    if(snapshot.version!==this.version||this.paragraphSnapshots[snapshot.index]!==snapshot||this.opcPackage.text(DOCUMENT_PART)!==this.xml)fail('docx-stale-paragraph','DOCX paragraph formatting handle is stale');
  }

  private assertFormattingUnprotected(): void {
    for(const rel of this.opcPackage.relationships(DOCUMENT_PART)){
      if(rel.type!=='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings')continue;
      if(rel.external)fail('docx-format-protected','External settings cannot be checked');
      const settings=parseXml(this.opcPackage.text(rel.resolved!));
      if(!isWord(settings.root,'settings'))fail('docx-format-protected','Invalid settings root');
      for(const guard of elements(settings,'documentProtection',W_NS))if(!/^(0|false|off)$/i.test(attribute(guard,'enforcement',W_NS)??''))fail('docx-format-protected','Document protection refuses formatting');
    }
  }

  private commitParagraphProperties(nextXml:string):void {
    const nextVersion=this.version+1,nextDocument=parseXml(nextXml);
    const collections=collectDocumentCollections(nextDocument,nextVersion,this.tableVersion);
    if(JSON.stringify(collections.paragraphs.paragraphs.map(p=>p.text))!==JSON.stringify(this.paragraphSnapshots.map(p=>p.text)))fail('docx-format-unsafe','Formatting unexpectedly changed paragraph text');
    this.opcPackage.transaction(()=>{this.opcPackage.set(DOCUMENT_PART,nextXml);this.opcPackage.toBytes();});
    this.xml=nextXml;this.version=nextVersion;this.xmlDocument=nextDocument;this.applyDocumentCollections(collections);
  }

  private assertRowHeaderTarget(snapshot:TableSnapshot,row:number):void {
    if(snapshot.version!==this.tableVersion||this.tableSnapshots[snapshot.index]!==snapshot)fail('docx-stale-table','Table snapshot is stale');
    if(this.opcPackage.text(DOCUMENT_PART)!==this.xml)fail('docx-stale-table','Document XML changed outside the table handle');
    if(!Number.isInteger(row)||row<0||row>=snapshot.rows)throw new RangeError('DOCX table row is out of range');
    if(snapshot.unsupported)fail('docx-table-unsupported',`Table contains unsupported topology: ${snapshot.unsupported}`);
  }

  getRowHeader(snapshot:TableSnapshot,row:number):boolean {
    this.assertRowHeaderTarget(snapshot,row);
    return readRowHeader(this.xml,snapshot.element,row);
  }

  setRowHeader(snapshot:TableSnapshot,row:number,value:boolean|null):{changed:number} {
    this.assertRowHeaderTarget(snapshot,row);this.assertFormattingUnprotected();
    const next=editRowHeader(this.xml,snapshot.element,row,value);if(next===this.xml)return {changed:0};
    const tables=collectTables(parseXml(next),this.tableVersion,this.version+1);
    const grids=(items:TableSnapshot[])=>items.map(t=>[t.rows,t.columns]);
    if(JSON.stringify(grids(tables.tables))!==JSON.stringify(grids(this.tableSnapshots)))fail('docx-row-header-unsafe','Row-header edit changed table dimensions');
    this.commitParagraphProperties(next);return {changed:1};
  }

  resolveTable(table: TableSnapshot): TableSnapshot {
    if (this.tableVersion !== table.version) {
      fail("docx-stale-table", "DOCX table handle is stale after document mutation");
    }

    const liveTable = this.tableSnapshots[table.index];
    if (!liveTable) {
      fail("docx-table-missing", `Missing DOCX table ${table.index + 1}`);
    }
    return liveTable;
  }

  private checkedCellProperties(cell:TableCellSnapshot):TableCellSnapshot {
    if(this.version!==cell.version||this.opcPackage.text(DOCUMENT_PART)!==this.xml)fail('docx-stale-table-cell','Cell properties target is stale');
    const live=this.resolveTableCell(cell);
    if(live!==cell)fail('docx-stale-table-cell','Cell properties target is not the current snapshot');
    if(live.unsupported)fail('docx-table-cell-unsupported','Cell properties require supported text topology');
    return live;
  }

  inspectCellProperties(cell:TableCellSnapshot):DirectCellProperties {
    return readCellProperties(this.xml,this.checkedCellProperties(cell).element);
  }

  setCellProperties(cell:TableCellSnapshot,input:CellPropertiesPatch):{changed:number} {
    const live=this.checkedCellProperties(cell),patch=normalizeCellProperties(input);
    this.assertFormattingUnprotected();
    const next=editCellProperties(this.xml,live.element,patch);
    if(next===this.xml)return {changed:0};
    this.commitParagraphProperties(next);return {changed:1};
  }

  replaceTableCellText(cell: TableCellSnapshot, text: string): void {
    if (this.version !== cell.version) {
      fail("docx-stale-table-cell", "DOCX table cell handle is stale after document mutation");
    }
    if (typeof text !== "string") {
      fail("docx-invalid-argument", "DOCX table cell text must be a string");
    }

    const liveCell = this.resolveTableCell(cell);
    if (liveCell.unsupported) {
      fail(
        "docx-table-cell-unsupported",
        `DOCX table cell (${liveCell.row},${liveCell.column}) uses unsupported topology ${liveCell.unsupported}`,
      );
    }

    const tcPr = liveCell.element.children.find((child) => isWord(child, "tcPr"));
    const firstParagraph = liveCell.element.children.find((child) => isWord(child, "p"));
    const firstParagraphProperties = firstParagraph?.children.find((child) => isWord(child, "pPr"));
    const firstRun = firstParagraph?.children.find((child) => isWord(child, "r"));
    const firstRunProperties = firstRun?.children.find((child) => isWord(child, "rPr"));
    const replacement = [
      tcPr ? rawXml(this.xml, tcPr) : "",
      buildTableCellParagraphsXml(
        normalizeCellText(text),
        firstParagraphProperties ? rawXml(this.xml, firstParagraphProperties) : "",
        firstRunProperties ? rawXml(this.xml, firstRunProperties) : "",
        inheritedNamespaces(firstParagraph),
        inheritedNamespaces(firstRun),
      ),
    ].join("");

    const nextXml = applyEdits(this.xml, [{
      start: liveCell.element.openEnd,
      end: liveCell.element.closeStart,
      value: replacement,
    }]);
    const nextVersion = this.version + 1;
    const nextDocument = parseXml(nextXml);
    const nextCollections = collectDocumentCollections(nextDocument, nextVersion, this.tableVersion);

    this.opcPackage.transaction(() => {
      this.opcPackage.set(DOCUMENT_PART, nextXml);
    });

    this.xml = nextXml;
    this.version = nextVersion;
    this.xmlDocument = nextDocument;
    this.applyDocumentCollections(nextCollections);
  }

  private resolveTableCell(cell: TableCellSnapshot): TableCellSnapshot {
    const table = this.tableSnapshots[cell.tableIndex];
    if (!table) {
      fail("docx-table-missing", `Missing DOCX table ${cell.tableIndex + 1}`);
    }
    if (table.unsupported) {
      fail("docx-table-unsupported", `DOCX table ${cell.tableIndex + 1} uses unsupported topology ${table.unsupported}`);
    }

    const entry = table.cells[cell.row]?.[cell.column];
    if (!entry) {
      fail("docx-table-unsupported", `DOCX table ${cell.tableIndex + 1} cannot resolve cell (${cell.row},${cell.column}) safely`);
    }
    if (entry.kind === "merged") {
      fail("docx-table-merged-cell", `DOCX table ${cell.tableIndex + 1} cell (${cell.row},${cell.column}) is merged`);
    }
    return entry.cell;
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
    this.reloadStructures();
  }

  private reloadStructures(): void {
    this.xmlDocument = parseXml(this.xml);
    this.applyDocumentCollections(collectDocumentCollections(this.xmlDocument, this.version, this.tableVersion));
  }

  private applyDocumentCollections(collections: DocumentCollections): void {
    this.applyParagraphCollection(collections.paragraphs);
    this.applyTableCollection(collections.tables);
  }

  private applyParagraphCollection(collection: ParagraphCollection): void {
    this.omittedTopologies = collection.omittedTopologies;
    this.paragraphSnapshots = collection.paragraphs;
    this.paragraphHandles = this.paragraphSnapshots.map(
      (snapshot) => new Paragraph(this, snapshot),
    );
  }

  private applyTableCollection(collection: TableCollection): void {
    this.tableSnapshots = collection.tables;
    this.tableHandles = this.tableSnapshots.map(
      (snapshot) => new Table(this, snapshot),
    );
  }

  private requireBody(): XmlElement {
    const body = this.xmlDocument.root.children.find((child) => isWord(child, "body"));
    if (!body) {
      fail("docx-document-invalid", "DOCX document XML is missing w:body");
    }
    return body;
  }

  private assertSupportedParagraphStyle(styleId: string): void {
    const links=this.opcPackage.relationships(DOCUMENT_PART).filter(r=>r.type==='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');
    if(links.length>1||links.some(r=>r.external))fail('docx-style-unsupported','Style relationship must be unique and internal');
    const stylesPart = links[0]?.resolved;
    if (!stylesPart) {
      fail(
        "docx-style-unsupported",
        `DOCX paragraph style ` + styleId + ` requires a resolved styles part`,
      );
    }

    if(getContentType(this.opcPackage,stylesPart)!=='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml')fail('docx-style-part-invalid','Unexpected styles content type');
    const stylesDocument = parseXml(this.opcPackage.text(stylesPart));
    if (!isWord(stylesDocument.root, "styles")) {
      fail("docx-style-part-invalid", `DOCX styles part ` + stylesPart + ` is invalid`);
    }

    const selected=stylesDocument.root.children.filter(style=>isWord(style,'style')&&attribute(style,'styleId',W_NS)===styleId);
    if(selected.length>1)fail('docx-style-ambiguous','Duplicate style ID');
    if(selected.length!==1||attribute(selected[0]!,'type',W_NS)!=='paragraph') {
      fail("docx-style-missing", `Unknown DOCX paragraph style ` + styleId);
    }
  }
}

export async function open(input: OpenInput): Promise<Document> {
  return Document.open(input);
}

export function create(): Document {
  return Document.create();
}

export async function save(document: Document, path?: string): Promise<Uint8Array> {
  return document.save(path);
}

function collectDocumentCollections(document: XmlDocument, version: number, tableVersion: number): DocumentCollections {
  return {
    paragraphs: collectParagraphs(document, version),
    tables: collectTables(document, tableVersion, version),
  };
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

function collectTables(document: XmlDocument, tableVersion: number, cellVersion: number): TableCollection {
  const body = document.root.children.find((child) => isWord(child, "body"));
  if (!body) {
    fail("docx-document-invalid", "DOCX document XML is missing w:body");
  }

  const tables: TableSnapshot[] = [];
  for (const child of body.children) {
    if (isWord(child, "tbl")) {
      tables.push(analyzeTopLevelTable(child, tables.length, tableVersion, cellVersion));
    }
  }

  return { tables };
}

function analyzeTopLevelTable(table: XmlElement, index: number, tableVersion: number, cellVersion: number): TableSnapshot {
  const rowElements = table.children.filter((child) => isWord(child, "tr"));
  const columns = readTableColumnCount(table);
  const cells: Array<Array<TableGridEntry | undefined>> = [];
  let unsupported: string | undefined;

  for (const child of table.children) {
    if (isWord(child, "tblPr") || isWord(child, "tblGrid") || isWord(child, "tblPrEx") || isWord(child, "tr")) {
      continue;
    }
    unsupported = unsupported ?? child.name;
  }

  for (const [rowIndex, row] of rowElements.entries()) {
    const rowCells: Array<TableGridEntry | undefined> = Array.from({ length: columns }, () => undefined);
    let cursor = 0;

    for (const child of row.children) {
      if (!isWord(child, "trPr") && !isWord(child, "tc")) {
        unsupported = unsupported ?? child.name;
      }
    }

    const rowProperties = row.children.find((child) => isWord(child, "trPr"));
    const gridBefore = readOptionalVal(rowProperties, "gridBefore");
    const gridAfter = readOptionalVal(rowProperties, "gridAfter");
    if (gridBefore !== 0 || gridAfter !== 0) {
      unsupported = unsupported ?? (gridBefore !== 0 ? "w:gridBefore" : "w:gridAfter");
    }

    for (const cell of row.children) {
      if (!isWord(cell, "tc")) {
        continue;
      }

      const tcPr = cell.children.find((child) => isWord(child, "tcPr"));
      const gridSpan = readGridSpan(tcPr);
      if (gridSpan < 1) {
        unsupported = unsupported ?? "w:gridSpan";
        continue;
      }
      if (cursor + gridSpan > columns) {
        unsupported = unsupported ?? "row-width";
        break;
      }

      const merged = gridSpan > 1 || hasWordChild(tcPr, "vMerge") || hasWordChild(tcPr, "hMerge");
      if (merged) {
        for (let offset = 0; offset < gridSpan; offset += 1) {
          rowCells[cursor + offset] = { kind: "merged" };
        }
      } else {
        rowCells[cursor] = { kind: "cell", cell: analyzeTableCell(cell, cellVersion, index, rowIndex, cursor) };
      }
      cursor += gridSpan;
    }

    if (cursor !== columns) {
      unsupported = unsupported ?? "row-width";
    }
    cells.push(rowCells);
  }

  return {
    index,
    version: tableVersion,
    rows: rowElements.length,
    columns,
    element: table,
    cells,
    unsupported,
  };
}

function analyzeTableCell(
  cell: XmlElement,
  version: number,
  tableIndex: number,
  row: number,
  column: number,
): TableCellSnapshot {
  const paragraphs: TableParagraphSnapshot[] = [];
  let unsupported: string | undefined;

  for (const child of cell.children) {
    if (isWord(child, "tcPr")) {
      continue;
    }
    if (isWord(child, "tbl")) {
      unsupported = unsupported ?? child.name;
      continue;
    }
    if (!isWord(child, "p")) {
      unsupported = unsupported ?? child.name;
      continue;
    }

    const paragraph = analyzeParagraph(child, 0, version);
    paragraphs.push({
      text: paragraph.text,
      searchable: paragraph.searchable,
      unsupported: paragraph.unsupported,
    });
    if (!paragraph.searchable) {
      unsupported = unsupported ?? paragraph.unsupported;
    }
  }

  return {
    version,
    tableIndex,
    row,
    column,
    element: cell,
    paragraphs,
    unsupported,
  };
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
    element: paragraph,
    index,
    version,
    text: textParts.join(""),
    segments,
    searchable: unsupported === undefined,
    unsupported,
  };
}

function normalizeAddParagraphOptions(options: AddParagraphOptions | undefined): NormalizedAddParagraphOptions {
  if (options === undefined) {
    return { bold: false, italic: false };
  }
  if (options === null || typeof options !== "object" || Array.isArray(options)) {
    fail("docx-invalid-argument", "DOCX paragraph options must be an object when provided");
  }

  const allowedKeys = new Set(["bold", "italic", "style"]);
  for (const key of Object.keys(options)) {
    if (!allowedKeys.has(key)) {
      fail("docx-invalid-argument", `DOCX paragraph option ${key} is unsupported`);
    }
  }

  if (options.bold !== undefined && typeof options.bold !== "boolean") {
    fail("docx-invalid-argument", "DOCX paragraph bold option must be boolean");
  }
  if (options.italic !== undefined && typeof options.italic !== "boolean") {
    fail("docx-invalid-argument", "DOCX paragraph italic option must be boolean");
  }
  if (options.style !== undefined && typeof options.style !== "string") {
    fail("docx-invalid-argument", "DOCX paragraph style option must be a string");
  }
  if (options.style === "") {
    fail("docx-invalid-argument", "DOCX paragraph style option must not be empty");
  }

  return {
    bold: options.bold ?? false,
    italic: options.italic ?? false,
    style: options.style,
  };
}

function buildParagraphXml(text: string, options: NormalizedAddParagraphOptions): string {
  const paragraphProperties = options.style
    ? `<w:pPr><w:pStyle w:val="${escapeAttribute(options.style)}"/></w:pPr>`
    : "";
  const runProperties = [
    options.bold ? "<w:b/>" : "",
    options.italic ? "<w:i/>" : "",
  ].join("");
  const textAttributes = needsPreserveSpace(text) ? " xml:space=\"preserve\"" : "";
  const run = text.length > 0 || runProperties.length > 0
    ? `<w:r>${runProperties ? `<w:rPr>${runProperties}</w:rPr>` : ""}<w:t${textAttributes}>${escapeText(text)}</w:t></w:r>`
    : "";
  return `<w:p xmlns:w="${W_NS}">${paragraphProperties}${run}</w:p>`;
}

function buildTableXml(rows: number, columns: number): string {
  const cellWidth = Math.max(1, Math.floor(DEFAULT_TABLE_WIDTH_DXA / columns));
  const gridColumns = Array.from({ length: columns }, () => `<w:gridCol w:w="${cellWidth}"/>`).join("");
  const rowXml = `<w:tr>${Array.from({ length: columns }, () => (
    `<w:tc><w:tcPr><w:tcW w:w="${cellWidth}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`
  )).join("")}</w:tr>`;

  return [
    `<w:tbl xmlns:w="${W_NS}">`,
    "<w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>",
    `<w:tblGrid>${gridColumns}</w:tblGrid>`,
    Array.from({ length: rows }, () => rowXml).join(""),
    "</w:tbl>",
  ].join("");
}

function buildTableCellParagraphsXml(text: string, firstParagraphPropertiesXml: string, firstRunPropertiesXml: string, paragraphNamespaces:Record<string,string>,runNamespaces:Record<string,string>): string {
  return text.split("\n").map((line, index) =>
    buildTableCellParagraphXml(
      line,
      index === 0 ? firstParagraphPropertiesXml : "",
      index === 0 ? firstRunPropertiesXml : "",
      paragraphNamespaces,
      runNamespaces,
    )
  ).join("");
}

function buildTableCellParagraphXml(text: string, paragraphPropertiesXml: string, runPropertiesXml: string,paragraphNamespaces:Record<string,string>,runNamespaces:Record<string,string>): string {
  const textAttributes = needsPreserveSpace(text) ? " xml:space=\"preserve\"" : "";
  const run = text.length > 0 || runPropertiesXml.length > 0
    ? `<w:r${namespaceAttributes(runNamespaces)}>${runPropertiesXml}<w:t${textAttributes}>${escapeText(text)}</w:t></w:r>`
    : "";
  return `<w:p${namespaceAttributes(paragraphNamespaces)}>${paragraphPropertiesXml}${run}</w:p>`;
}

function inheritedNamespaces(node:XmlElement|undefined):Record<string,string>{
  const chain:XmlElement[]=[];for(let current=node;current;current=current.parent)chain.push(current);
  const result:Record<string,string>=Object.create(null);
  for(const ancestor of chain.reverse())for(const [name,value]of Object.entries(ancestor.attributes))if(name==='xmlns'||name.startsWith('xmlns:'))result[name]=value;
  // Authored w elements must have the Word namespace. A conflicting lexical w
  // in copied formatting cannot be reinterpreted safely; refuse before mutation.
  if(result['xmlns:w']&&result['xmlns:w']!==W_NS)fail('docx-table-cell-unsupported','Conflicting w namespace in copied formatting');
  result['xmlns:w']=W_NS;return result;
}
function namespaceAttributes(map:Record<string,string>):string{return Object.entries(map).map(([name,value])=>` ${name}="${escapeAttribute(value)}"`).join('');}

function buildMinimalDocxParts(): Map<string, Uint8Array> {
  return new Map([
    ["[Content_Types].xml", UTF8_ENCODER.encode(minimalContentTypesXml())],
    ["_rels/.rels", UTF8_ENCODER.encode(minimalRootRelationshipsXml())],
    [DOCUMENT_PART, UTF8_ENCODER.encode(minimalDocumentXml())],
  ]);
}

function minimalContentTypesXml(): string {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    `<Types xmlns="${CONTENT_TYPES_NS}">`,
    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
    `<Override PartName="/word/document.xml" ContentType="${DOCUMENT_CONTENT_TYPE}"/>`,
    "</Types>",
  ].join("");
}

function minimalRootRelationshipsXml(): string {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    `<Relationships xmlns="${REL_NS}">`,
    `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="word/document.xml"/>`,
    "</Relationships>",
  ].join("");
}

function minimalDocumentXml(): string {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    `<w:document xmlns:w="${W_NS}">`,
    "<w:body>",
    "<w:sectPr>",
    "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>",
    "<w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>",
    "<w:cols w:space=\"720\"/>",
    "<w:docGrid w:linePitch=\"360\"/>",
    "</w:sectPr>",
    "</w:body>",
    "</w:document>",
  ].join("");
}

function assertTableDimensions(rows: number, columns: number): void {
  if (!Number.isInteger(rows) || rows < 1 || rows > MAX_TABLE_ROWS) {
    throw new RangeError(`DOCX table rows must be an integer between 1 and ${MAX_TABLE_ROWS}`);
  }
  if (!Number.isInteger(columns) || columns < 1 || columns > MAX_TABLE_COLUMNS) {
    throw new RangeError(`DOCX table columns must be an integer between 1 and ${MAX_TABLE_COLUMNS}`);
  }
  if (rows * columns > MAX_TABLE_CELLS) {
    throw new RangeError(`DOCX table area must not exceed ${MAX_TABLE_CELLS} cells`);
  }
}

function assertTableCoordinate(table: TableSnapshot, row: number, column: number): void {
  if (!Number.isInteger(row) || row < 0 || row >= table.rows) {
    throw new RangeError(`DOCX table row ${row} is out of range`);
  }
  if (!Number.isInteger(column) || column < 0 || column >= table.columns) {
    throw new RangeError(`DOCX table column ${column} is out of range`);
  }
}

function readTableColumnCount(table: XmlElement): number {
  const grid = table.children.find((child) => isWord(child, "tblGrid"));
  if (grid) {
    return grid.children.filter((child) => isWord(child, "gridCol")).length;
  }

  let columns = 0;
  for (const row of table.children) {
    if (!isWord(row, "tr")) {
      continue;
    }
    let width = 0;
    for (const child of row.children) {
      if (!isWord(child, "tc")) {
        continue;
      }
      const tcPr = child.children.find((grandchild) => isWord(grandchild, "tcPr"));
      width += readGridSpan(tcPr);
    }
    columns = Math.max(columns, width);
  }
  return columns;
}

function readGridSpan(tcPr: XmlElement | undefined): number {
  const gridSpan = tcPr?.children.find((child) => isWord(child, "gridSpan"));
  if (!gridSpan) {
    return 1;
  }
  const value = attribute(gridSpan, "val", W_NS);
  if (value === undefined) {
    return 0;
  }
  const parsed = Number.parseInt(value, 10);
  return Number.isInteger(parsed) ? parsed : 0;
}

function readOptionalVal(container: XmlElement | undefined, childName: string): number {
  const child = container?.children.find((candidate) => isWord(candidate, childName));
  if (!child) {
    return 0;
  }
  const value = attribute(child, "val", W_NS);
  if (value === undefined) {
    return 1;
  }
  const parsed = Number.parseInt(value, 10);
  return Number.isInteger(parsed) ? parsed : 1;
}

function hasWordChild(container: XmlElement | undefined, localName: string): boolean {
  return container?.children.some((child) => isWord(child, localName)) ?? false;
}

function rawXml(xml: string, element: XmlElement): string {
  return xml.slice(element.start, element.end);
}

function normalizeCellText(text: string): string {
  return text.replace(/\r\n?/g, "\n");
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

export { inspectStories, storyParts, type RevisionView, type Story, type StoryInspection, type StoryKind } from "./story.ts";
export { inspectRevisions, resolveRevisions, type Revision, type RevisionFinding } from "./revisions.ts";
// Review mutations operate on OpcPackage snapshots; reopen Document afterward.
export { trackedReplace } from "./redline.ts";
export { inspectComments, setCommentResolved, type CommentInfo, type CommentFinding, type CommentInspection } from "./comments.ts";
