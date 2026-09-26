import { posix } from "node:path";

import { OoxmlError } from "../errors.ts";
import {appendTextBox,textBoxRequest,type TextBoxGeometry,type TextBoxOptions,type TextBoxReceipt} from './text-box.ts';
export type {TextBoxGeometry,TextBoxOptions,TextBoxReceipt} from './text-box.ts';
import { addPart, addRelationship, nextPartName } from "../opc/graph.ts";
import { OpcPackage, type Relationship } from "../opc/package.ts";
import { attribute, applyEdits, elements, escapeAttribute, escapeText, parseXml, type XmlElement } from "../xml/index.ts";

const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const TABLE_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";
const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

const OFFICE_DOCUMENT_RELATIONSHIP = `${OFFICE_REL_NS}/officeDocument`;
const SLIDE_RELATIONSHIP = `${OFFICE_REL_NS}/slide`;
const SLIDE_MASTER_RELATIONSHIP = `${OFFICE_REL_NS}/slideMaster`;
const SLIDE_LAYOUT_RELATIONSHIP = `${OFFICE_REL_NS}/slideLayout`;
const THEME_RELATIONSHIP = `${OFFICE_REL_NS}/theme`;
const VIEW_PROPS_RELATIONSHIP = `${OFFICE_REL_NS}/viewProps`;
const PRES_PROPS_RELATIONSHIP = `${OFFICE_REL_NS}/presProps`;
const TABLE_STYLES_RELATIONSHIP = `${OFFICE_REL_NS}/tableStyles`;

const PRESENTATION_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const PRES_PROPS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
const VIEW_PROPS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
const TABLE_STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
const THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";
const SLIDE_MASTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
const SLIDE_LAYOUT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

const DEFAULT_SLIDE_WIDTH = "9144000";
const DEFAULT_SLIDE_HEIGHT = "6858000";
const DEFAULT_NOTES_WIDTH = "6858000";
const DEFAULT_NOTES_HEIGHT = "9144000";
const DEFAULT_TABLE_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const MAX_PPTX_SHAPE_ID = 2147483647;
const MAX_TABLE_CELLS = 10_000;
const encoder = new TextEncoder();

type OpenInput = string | Uint8Array | ArrayBuffer;

type StoryFragment = {
  text: string;
  attrs: Record<string, string>;
};

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
  fragments: StoryFragment[];
  editableRuns: StoryRun[];
  readable: boolean;
  replaceable: boolean;
};

type PlaceholderKind = "title" | "subtitle";

type PlaceholderDescriptor = {
  type: string;
  idx?: string;
};

type SafeTextSlideLayout = {
  partName: string;
  title: PlaceholderDescriptor;
  subtitle?: PlaceholderDescriptor;
};

export type TableGeometry = {
  x: number;
  y: number;
  width: number;
  height: number;
};

type ResolvedSlideTable = {
  shapeId: number;
  table: XmlElement;
  gridColumns: XmlElement[];
  rows: XmlElement[];
  cells: XmlElement[][];
  merged: boolean;
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
 * Minimal PPTX read/write slice for relationship-ordered slide access, anchored text edits,
 * and conservative title-slide creation.
 *
 * Current authoring limit: existing decks accept `addTextSlide()` only when exactly one direct,
 * title-layout-compatible slide layout can be selected safely. This slice does not guess among
 * several plausible layouts or attempt broader inheritance synthesis.
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

  static create(): Presentation {
    return new Presentation(OpcPackage.fromParts(createBlankPresentationParts()));
  }

  static async open(input: OpenInput): Promise<Presentation> {
    const pkg = await OpcPackage.open(
      input instanceof ArrayBuffer ? new Uint8Array(input.slice(0)) : input,
    );
    return new Presentation(pkg);
  }

  get slideCount(): number {
    return this.slides.length;
  }

  get slideMasterCount(): number {
    return new Set(resolveSlideMasterParts(this.package, this.mainPartName)).size;
  }

  get slideLayoutCount(): number {
    return new Set(resolveSlideLayoutParts(this.package, this.mainPartName)).size;
  }

  /**
   * Appends a title slide using a compatible, non-ambiguous layout only.
   *
   * New decks created by `Presentation.create()` carry one owned title layout. Existing decks
   * are edited only when exactly one direct title-slide layout is discoverable; otherwise this
   * method refuses before mutating package bytes.
   */
  addTextSlide(title: string, subtitle?: string): Slide {
    validateTextSlideArgs(title, subtitle);

    let createdPartName = "";
    this.package.transaction(() => {
      const layout = chooseSafeTextSlideLayout(this.package, this.mainPartName, subtitle !== undefined);
      createdPartName = nextPartName(this.package, "ppt/slides/slide%d.xml");

      addXmlPart(
        this.package,
        createdPartName,
        buildTextSlideXml(layout, title, subtitle),
        SLIDE_CONTENT_TYPE,
      );
      addInternalRelationship(
        this.package,
        createdPartName,
        SLIDE_LAYOUT_RELATIONSHIP,
        relativeTarget(createdPartName, layout.partName),
      );

      const slideRelationship = addInternalRelationship(
        this.package,
        this.mainPartName,
        SLIDE_RELATIONSHIP,
        relativeTarget(this.mainPartName, createdPartName),
      );
      const nextXml = appendSlideId(this.package.text(this.mainPartName), nextSlideId(this.package, this.mainPartName), slideRelationship.id);
      this.package.set(this.mainPartName, nextXml);
      this.package.toBytes();
    });

    const slide = new Slide(this, createdPartName, this.slides.length);
    this.slideVersions.set(createdPartName, 0);
    this.slides.push(slide);
    return slide;
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

  get tables(): Table[] {
    const version = this.presentation.currentSlideVersion(this.partName);
    return collectSlideTables(this.presentation.package.text(this.partName), this.partName)
      .map((table, tableIndex) => new Table(this, tableIndex, version, table.shapeId));
  }

  /** Append a slide-space text box. Newline sequences become separate paragraphs. */
  addTextBox(text: string, geometry: TextBoxGeometry, options: TextBoxOptions = {}): TextBoxReceipt {
    const request=textBoxRequest(text,geometry,options),pkg=this.presentation.package;
    const main=pkg.mainPart();
    if(resolveSlideParts(pkg,main)[this.index]!==this.partName)throw new OoxmlError('PPTX_STALE_SLIDE','Slide order or identity changed outside this handle');
    if(elements(parseXml(pkg.text(main)),'modifyVerifier',PRESENTATION_NS).length)throw new OoxmlError('PPTX_PROTECTED','Presentation modification protection refuses text-box authoring');
    const result=appendTextBox(pkg.text(this.partName),request);
    pkg.transaction(()=>{pkg.set(this.partName,result.xml);pkg.toBytes();});
    this.presentation.bumpSlideVersion(this.partName);
    return {shapeId:result.shapeId,partName:this.partName,paragraphCount:result.paragraphCount};
  }

  addTable(rows: number, columns: number, geometry: TableGeometry): Table {
    validateAddTableArgs(rows, columns, geometry);

    let tableIndex = 0;
    this.presentation.package.transaction(() => {
      const xml = this.presentation.package.text(this.partName);
      const document = parseXml(xml);
      const spTree = requireSlideShapeTree(document, this.partName);
      tableIndex = collectSlideTablesFromShapeTree(spTree, this.partName).length;
      const shapeId = nextSlideShapeId(document, this.partName);
      const nextXml = insertShapeTreeChild(
        xml,
        spTree,
        buildTableGraphicFrameXml(shapeId, rows, columns, geometry),
      );
      this.presentation.package.set(this.partName, nextXml);
      this.presentation.package.toBytes();
    });

    this.presentation.bumpSlideVersion(this.partName);
    const table = this.tables[tableIndex];
    if (!table) {
      throw new OoxmlError(
        "PPTX_SLIDE_INVALID",
        `Authored table could not be resolved in ${this.partName}`,
      );
    }
    return table;
  }

  inspectText(_label: string): InspectedParagraph[] {
    const xml = this.presentation.package.text(this.partName);
    const paragraphs = collectStoryParagraphs(xml);
    assertReadableParagraphs(this.partName, paragraphs);
    const version = this.presentation.currentSlideVersion(this.partName);

    return paragraphs.map((paragraph, paragraphIndex) => ({
      text: paragraph.text,
      runs: paragraph.fragments.map((fragment) => ({ text: fragment.text, attrs: { ...fragment.attrs } })),
      anchor: {
        kind: "pptx-text",
        part: this.partName,
        paragraphIndex,
        version,
        text: paragraph.text,
      },
    }));
  }

  resolveTableHandle(tableIndex: number, version: number, shapeId: number): ResolvedSlideTable {
    if (version !== this.presentation.currentSlideVersion(this.partName)) {
      throw staleTableHandle(this.partName, "table handle is stale after slide mutation");
    }

    const table = collectSlideTables(this.presentation.package.text(this.partName), this.partName)[tableIndex];
    if (!table || table.shapeId !== shapeId) {
      throw staleTableHandle(this.partName, "table identity has changed");
    }
    return table;
  }

  setTableCellText(
    tableIndex: number,
    version: number,
    shapeId: number,
    row: number,
    column: number,
    value: string,
  ): void {
    if (typeof value !== "string") {
      throw new OoxmlError("PPTX_ARGUMENT_INVALID", "Table cell text must be a string");
    }

    let changed = false;
    this.presentation.package.transaction(() => {
      if (version !== this.presentation.currentSlideVersion(this.partName)) {
        throw staleTableHandle(this.partName, "table handle is stale after slide mutation");
      }

      const xml = this.presentation.package.text(this.partName);
      const table = collectSlideTables(xml, this.partName)[tableIndex];
      if (!table || table.shapeId !== shapeId) {
        throw staleTableHandle(this.partName, "table identity has changed");
      }
      if (table.merged) {
        throw mergedTableUnsupported(this.partName, row, column);
      }

      const cell = table.cells[row]?.[column];
      if (!cell) {
        throw staleTableHandle(this.partName, "table grid has changed");
      }

      const nextXml = replaceTableCellText(xml, cell, this.partName, row, column, value);
      this.presentation.package.set(this.partName, nextXml);
      this.presentation.package.toBytes();
      changed = nextXml !== xml;
    });

    if (changed) {
      this.presentation.bumpSlideVersion(this.partName);
    }
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

    const paragraphs = collectStoryParagraphs(this.presentation.package.text(notesPart));
    assertReadableParagraphs(notesPart, paragraphs);
    return paragraphs.map((paragraph) => paragraph.text).join("\n");
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

export class Table {
  constructor(
    private readonly slideRef: Slide,
    private readonly tableIndex: number,
    private readonly version: number,
    private readonly shapeId: number,
  ) {}

  get rows(): number {
    return this.resolve().rows.length;
  }

  get columns(): number {
    return this.resolve().gridColumns.length;
  }

  cell(row: number, column: number): TableCell {
    const table = this.resolve();
    assertTableIndex(row, table.rows.length, "row");
    assertTableIndex(column, table.gridColumns.length, "column");
    return new TableCell(this, row, column);
  }

  getCellText(row: number, column: number): string {
    const table = this.resolve();
    assertTableIndex(row, table.rows.length, "row");
    assertTableIndex(column, table.gridColumns.length, "column");
    return readTableCellText(table.cells[row]![column]!, this.slideRef.partName, row, column);
  }

  setCellText(row: number, column: number, value: string): void {
    const table = this.resolve();
    assertTableIndex(row, table.rows.length, "row");
    assertTableIndex(column, table.gridColumns.length, "column");
    this.slideRef.setTableCellText(this.tableIndex, this.version, this.shapeId, row, column, value);
  }

  private resolve(): ResolvedSlideTable {
    return this.slideRef.resolveTableHandle(this.tableIndex, this.version, this.shapeId);
  }
}

export class TableCell {
  constructor(
    private readonly tableRef: Table,
    readonly row: number,
    readonly column: number,
  ) {}

  get text(): string {
    return this.tableRef.getCellText(this.row, this.column);
  }

  set text(value: string) {
    this.tableRef.setCellText(this.row, this.column, value);
  }
}

function createBlankPresentationParts(): Map<string, Uint8Array> {
  return new Map<string, Uint8Array>([
    ["[Content_Types].xml", xmlBytes(buildBlankContentTypesXml())],
    ["_rels/.rels", xmlBytes(buildRootRelationshipsXml())],
    ["ppt/presentation.xml", xmlBytes(buildBlankPresentationXml())],
    ["ppt/_rels/presentation.xml.rels", xmlBytes(buildBlankPresentationRelationshipsXml())],
    ["ppt/presProps.xml", xmlBytes(buildPresPropsXml())],
    ["ppt/viewProps.xml", xmlBytes(buildViewPropsXml())],
    ["ppt/tableStyles.xml", xmlBytes(buildTableStylesXml())],
    ["ppt/theme/theme1.xml", xmlBytes(buildThemeXml())],
    ["ppt/slideMasters/slideMaster1.xml", xmlBytes(buildSlideMasterXml())],
    ["ppt/slideMasters/_rels/slideMaster1.xml.rels", xmlBytes(buildSlideMasterRelationshipsXml())],
    ["ppt/slideLayouts/slideLayout1.xml", xmlBytes(buildTitleSlideLayoutXml())],
    ["ppt/slideLayouts/_rels/slideLayout1.xml.rels", xmlBytes(buildTitleSlideLayoutRelationshipsXml())],
  ]);
}

function buildBlankContentTypesXml(): string {
  return [
    xmlDeclaration(),
    `<Types xmlns="${CONTENT_TYPES_NS}">`,
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
    `<Override PartName="/ppt/presentation.xml" ContentType="${PRESENTATION_CONTENT_TYPE}"/>`,
    `<Override PartName="/ppt/presProps.xml" ContentType="${PRES_PROPS_CONTENT_TYPE}"/>`,
    `<Override PartName="/ppt/viewProps.xml" ContentType="${VIEW_PROPS_CONTENT_TYPE}"/>`,
    `<Override PartName="/ppt/tableStyles.xml" ContentType="${TABLE_STYLES_CONTENT_TYPE}"/>`,
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="${THEME_CONTENT_TYPE}"/>`,
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${SLIDE_MASTER_CONTENT_TYPE}"/>`,
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${SLIDE_LAYOUT_CONTENT_TYPE}"/>`,
    `</Types>`,
  ].join("");
}

function buildRootRelationshipsXml(): string {
  return [
    xmlDeclaration(),
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`,
    `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="ppt/presentation.xml"/>`,
    `</Relationships>`,
  ].join("");
}

function buildBlankPresentationXml(): string {
  return [
    xmlDeclaration(),
    `<p:presentation xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:p="${PRESENTATION_NS}" saveSubsetFonts="1" autoCompressPictures="0">`,
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>`,
    `<p:sldIdLst></p:sldIdLst>`,
    `<p:sldSz cx="${DEFAULT_SLIDE_WIDTH}" cy="${DEFAULT_SLIDE_HEIGHT}" type="screen4x3"/>`,
    `<p:notesSz cx="${DEFAULT_NOTES_WIDTH}" cy="${DEFAULT_NOTES_HEIGHT}"/>`,
    `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>`,
    `</p:presentation>`,
  ].join("");
}

function buildBlankPresentationRelationshipsXml(): string {
  return [
    xmlDeclaration(),
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`,
    `<Relationship Id="rId1" Type="${SLIDE_MASTER_RELATIONSHIP}" Target="slideMasters/slideMaster1.xml"/>`,
    `<Relationship Id="rId2" Type="${PRES_PROPS_RELATIONSHIP}" Target="presProps.xml"/>`,
    `<Relationship Id="rId3" Type="${VIEW_PROPS_RELATIONSHIP}" Target="viewProps.xml"/>`,
    `<Relationship Id="rId4" Type="${THEME_RELATIONSHIP}" Target="theme/theme1.xml"/>`,
    `<Relationship Id="rId5" Type="${TABLE_STYLES_RELATIONSHIP}" Target="tableStyles.xml"/>`,
    `</Relationships>`,
  ].join("");
}

function buildPresPropsXml(): string {
  return [
    xmlDeclaration(),
    `<p:presentationPr xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:p="${PRESENTATION_NS}"/>`,
  ].join("");
}

function buildViewPropsXml(): string {
  return [
    xmlDeclaration(),
    `<p:viewPr xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:p="${PRESENTATION_NS}">`,
    `<p:normalViewPr/>`,
    `<p:slideViewPr><p:cSldViewPr/></p:slideViewPr>`,
    `<p:notesTextViewPr/>`,
    `<p:gridSpacing cx="76200" cy="76200"/>`,
    `</p:viewPr>`,
  ].join("");
}

function buildTableStylesXml(): string {
  return [
    xmlDeclaration(),
    `<a:tblStyleLst xmlns:a="${DRAWING_NS}" def="${DEFAULT_TABLE_STYLE_ID}"/>`,
  ].join("");
}

function buildThemeXml(): string {
  return [
    xmlDeclaration(),
    `<a:theme xmlns:a="${DRAWING_NS}" name="Bun OOXML Theme">`,
    `<a:themeElements>`,
    `<a:clrScheme name="Bun">`,
    `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>`,
    `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>`,
    `<a:dk2><a:srgbClr val="1F497D"/></a:dk2>`,
    `<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>`,
    `<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>`,
    `<a:accent2><a:srgbClr val="C0504D"/></a:accent2>`,
    `<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>`,
    `<a:accent4><a:srgbClr val="8064A2"/></a:accent4>`,
    `<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>`,
    `<a:accent6><a:srgbClr val="F79646"/></a:accent6>`,
    `<a:hlink><a:srgbClr val="0000FF"/></a:hlink>`,
    `<a:folHlink><a:srgbClr val="800080"/></a:folHlink>`,
    `</a:clrScheme>`,
    `<a:fontScheme name="Bun">`,
    `<a:majorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>`,
    `<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>`,
    `</a:fontScheme>`,
    `<a:fmtScheme name="Bun">`,
    `<a:fillStyleLst>`,
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`,
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`,
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`,
    `</a:fillStyleLst>`,
    `<a:lnStyleLst>`,
    `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`,
    `<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`,
    `<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`,
    `</a:lnStyleLst>`,
    `<a:effectStyleLst>`,
    `<a:effectStyle><a:effectLst/></a:effectStyle>`,
    `<a:effectStyle><a:effectLst/></a:effectStyle>`,
    `<a:effectStyle><a:effectLst/></a:effectStyle>`,
    `</a:effectStyleLst>`,
    `<a:bgFillStyleLst>`,
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`,
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill>`,
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`,
    `</a:bgFillStyleLst>`,
    `</a:fmtScheme>`,
    `</a:themeElements>`,
    `</a:theme>`,
  ].join("");
}

function buildSlideMasterXml(): string {
  return [
    xmlDeclaration(),
    `<p:sldMaster xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:p="${PRESENTATION_NS}">`,
    `<p:cSld><p:spTree>${shapeTreePrefix()}</p:spTree></p:cSld>`,
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`,
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>`,
    `<p:txStyles>`,
    `<p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>`,
    `<p:bodyStyle><a:lvl1pPr marL="0" algn="l"><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle>`,
    `<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:otherStyle>`,
    `</p:txStyles>`,
    `</p:sldMaster>`,
  ].join("");
}

function buildSlideMasterRelationshipsXml(): string {
  return [
    xmlDeclaration(),
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`,
    `<Relationship Id="rId1" Type="${SLIDE_LAYOUT_RELATIONSHIP}" Target="../slideLayouts/slideLayout1.xml"/>`,
    `<Relationship Id="rId2" Type="${THEME_RELATIONSHIP}" Target="../theme/theme1.xml"/>`,
    `</Relationships>`,
  ].join("");
}

function buildTitleSlideLayoutXml(): string {
  return [
    xmlDeclaration(),
    `<p:sldLayout xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:p="${PRESENTATION_NS}" type="title" preserve="1">`,
    `<p:cSld name="Bun Title Slide"><p:spTree>`,
    shapeTreePrefix(),
    buildLayoutPlaceholderShapeXml({ id: 2, name: "Title 1", placeholder: { type: "ctrTitle" }, text: "Click to add title" }),
    buildLayoutPlaceholderShapeXml({ id: 3, name: "Subtitle 2", placeholder: { type: "subTitle", idx: "1" }, text: "Click to add subtitle" }),
    `</p:spTree></p:cSld>`,
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`,
    `</p:sldLayout>`,
  ].join("");
}

function buildTitleSlideLayoutRelationshipsXml(): string {
  return [
    xmlDeclaration(),
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`,
    `<Relationship Id="rId1" Type="${SLIDE_MASTER_RELATIONSHIP}" Target="../slideMasters/slideMaster1.xml"/>`,
    `</Relationships>`,
  ].join("");
}

function buildTextSlideXml(layout: SafeTextSlideLayout, title: string, subtitle?: string): string {
  const parts = [
    xmlDeclaration(),
    `<p:sld xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:p="${PRESENTATION_NS}">`,
    `<p:cSld><p:spTree>`,
    shapeTreePrefix(),
    buildSlidePlaceholderShapeXml({ id: 2, name: "Title 1", placeholder: layout.title, text: title }),
  ];

  if (subtitle !== undefined) {
    if (!layout.subtitle) {
      throw new OoxmlError("PPTX_LAYOUT_UNSAFE", `Layout ${layout.partName} is missing a direct subtitle placeholder`);
    }
    parts.push(buildSlidePlaceholderShapeXml({ id: 3, name: "Subtitle 2", placeholder: layout.subtitle, text: subtitle }));
  }

  parts.push(`</p:spTree></p:cSld>`, `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`, `</p:sld>`);
  return parts.join("");
}

function shapeTreePrefix(): string {
  return [
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`,
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`,
  ].join("");
}

function buildLayoutPlaceholderShapeXml(options: {
  id: number;
  name: string;
  placeholder: PlaceholderDescriptor;
  text: string;
}): string {
  return [
    `<p:sp>`,
    `<p:nvSpPr>`,
    `<p:cNvPr id="${options.id}" name="${escapeAttribute(options.name)}"/>`,
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>`,
    `<p:nvPr>${buildPlaceholderXml(options.placeholder)}</p:nvPr>`,
    `</p:nvSpPr>`,
    `<p:spPr/>`,
    `<p:txBody><a:bodyPr/><a:lstStyle/>${buildTextParagraphXml(options.text)}</p:txBody>`,
    `</p:sp>`,
  ].join("");
}

function buildSlidePlaceholderShapeXml(options: {
  id: number;
  name: string;
  placeholder: PlaceholderDescriptor;
  text: string;
}): string {
  return [
    `<p:sp>`,
    `<p:nvSpPr>`,
    `<p:cNvPr id="${options.id}" name="${escapeAttribute(options.name)}"/>`,
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>`,
    `<p:nvPr>${buildPlaceholderXml(options.placeholder)}</p:nvPr>`,
    `</p:nvSpPr>`,
    `<p:spPr/>`,
    `<p:txBody><a:bodyPr/><a:lstStyle/>${buildTextParagraphXml(options.text)}</p:txBody>`,
    `</p:sp>`,
  ].join("");
}

function buildPlaceholderXml(placeholder: PlaceholderDescriptor): string {
  const attrs = [`type="${escapeAttribute(placeholder.type)}"`];
  if (placeholder.idx !== undefined) {
    attrs.push(`idx="${escapeAttribute(placeholder.idx)}"`);
  }
  return `<p:ph ${attrs.join(" ")}/>`;
}

function buildTextParagraphXml(text: string): string {
  const preserve = needsPreserveSpace(text) ? ' xml:space="preserve"' : "";
  return `<a:p><a:r><a:t${preserve}>${escapeText(text)}</a:t></a:r></a:p>`;
}

function xmlDeclaration(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
}

function xmlBytes(xml: string): Uint8Array {
  return encoder.encode(xml);
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
      const relationshipId = attribute(slideId, "id", OFFICE_REL_NS);
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

function resolveSlideMasterParts(pkg: OpcPackage, mainPartName: string): string[] {
  const presentation = parseXml(pkg.text(mainPartName));
  if (presentation.root.localName !== "presentation" || presentation.root.namespaceURI !== PRESENTATION_NS) {
    throw new OoxmlError("PPTX_PRESENTATION_INVALID", `Invalid presentation root in ${mainPartName}`);
  }

  const masterList = presentation.root.children.find((child) => isElement(child, "sldMasterIdLst", PRESENTATION_NS));
  if (!masterList) {
    return [];
  }

  const relationships = new Map(
    pkg.relationships(mainPartName)
      .filter((relationship) => !relationship.external)
      .map((relationship) => [relationship.id, relationship]),
  );

  return masterList.children
    .filter((child) => isElement(child, "sldMasterId", PRESENTATION_NS))
    .map((masterId) => {
      const relationshipId = attribute(masterId, "id", OFFICE_REL_NS);
      const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
      if (!relationship || relationship.type !== SLIDE_MASTER_RELATIONSHIP || !relationship.resolved) {
        throw new OoxmlError(
          "PPTX_PRESENTATION_INVALID",
          `Slide master id in ${mainPartName} does not resolve to an internal slide master relationship`,
        );
      }
      return relationship.resolved;
    });
}

function resolveSlideLayoutParts(pkg: OpcPackage, mainPartName: string): string[] {
  return resolveSlideMasterParts(pkg, mainPartName).flatMap((partName) => resolveSlideLayoutPartsForMaster(pkg, partName));
}

function resolveSlideLayoutPartsForMaster(pkg: OpcPackage, masterPartName: string): string[] {
  const document = parseXml(pkg.text(masterPartName));
  if (document.root.localName !== "sldMaster" || document.root.namespaceURI !== PRESENTATION_NS) {
    throw new OoxmlError("PPTX_PRESENTATION_INVALID", `Invalid slide master root in ${masterPartName}`);
  }

  const layoutList = document.root.children.find((child) => isElement(child, "sldLayoutIdLst", PRESENTATION_NS));
  if (!layoutList) {
    return [];
  }

  const relationships = new Map(
    pkg.relationships(masterPartName)
      .filter((relationship) => !relationship.external)
      .map((relationship) => [relationship.id, relationship]),
  );

  return layoutList.children
    .filter((child) => isElement(child, "sldLayoutId", PRESENTATION_NS))
    .map((layoutId) => {
      const relationshipId = attribute(layoutId, "id", OFFICE_REL_NS);
      const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
      if (!relationship || relationship.type !== SLIDE_LAYOUT_RELATIONSHIP || !relationship.resolved) {
        throw new OoxmlError(
          "PPTX_PRESENTATION_INVALID",
          `Slide layout id in ${masterPartName} does not resolve to an internal slide layout relationship`,
        );
      }
      return relationship.resolved;
    });
}

function chooseSafeTextSlideLayout(
  pkg: OpcPackage,
  mainPartName: string,
  requireSubtitle: boolean,
): SafeTextSlideLayout {
  const matches = resolveSlideLayoutParts(pkg, mainPartName)
    .map((partName) => analyzeSafeTextSlideLayout(pkg, partName, requireSubtitle))
    .filter((layout): layout is SafeTextSlideLayout => layout !== undefined);

  if (matches.length === 1) {
    return matches[0]!;
  }

  if (matches.length === 0) {
    throw new OoxmlError(
      "PPTX_LAYOUT_UNSAFE",
      requireSubtitle
        ? "No compatible direct title/subtitle slide layout is available for safe PPTX authoring"
        : "No compatible direct title slide layout is available for safe PPTX authoring",
    );
  }

  throw new OoxmlError(
    "PPTX_LAYOUT_UNSAFE",
    `Several compatible title slide layouts are available (${matches.map((match) => match.partName).join(", ")}); refusing to guess`,
  );
}

function analyzeSafeTextSlideLayout(
  pkg: OpcPackage,
  partName: string,
  requireSubtitle: boolean,
): SafeTextSlideLayout | undefined {
  const document = parseXml(pkg.text(partName));
  if (document.root.localName !== "sldLayout" || document.root.namespaceURI !== PRESENTATION_NS) {
    return undefined;
  }
  if (document.root.attributes.type !== "title") {
    return undefined;
  }

  const cSld = document.root.children.find((child) => isElement(child, "cSld", PRESENTATION_NS));
  const spTree = cSld?.children.find((child) => isElement(child, "spTree", PRESENTATION_NS));
  if (!spTree) {
    return undefined;
  }

  const titles: PlaceholderDescriptor[] = [];
  const subtitles: PlaceholderDescriptor[] = [];

  for (const child of spTree.children) {
    if (!isElement(child, "sp", PRESENTATION_NS)) {
      continue;
    }
    const placeholder = directPlaceholder(child);
    if (!placeholder) {
      continue;
    }
    const type = placeholder.attributes.type;
    if (type === "title" || type === "ctrTitle") {
      titles.push({ type, idx: placeholder.attributes.idx });
      continue;
    }
    if (type === "subTitle") {
      subtitles.push({ type, idx: placeholder.attributes.idx });
    }
  }

  if (titles.length !== 1) {
    return undefined;
  }
  if (subtitles.length > 1) {
    return undefined;
  }
  if (requireSubtitle && subtitles.length !== 1) {
    return undefined;
  }

  return {
    partName,
    title: titles[0]!,
    subtitle: subtitles[0],
  };
}

function nextSlideId(pkg: OpcPackage, mainPartName: string): number {
  const presentation = parseXml(pkg.text(mainPartName));
  if (presentation.root.localName !== "presentation" || presentation.root.namespaceURI !== PRESENTATION_NS) {
    throw new OoxmlError("PPTX_PRESENTATION_INVALID", `Invalid presentation root in ${mainPartName}`);
  }

  const sldIdList = presentation.root.children.find((child) => isElement(child, "sldIdLst", PRESENTATION_NS));
  const maxId = sldIdList?.children
    .filter((child) => isElement(child, "sldId", PRESENTATION_NS))
    .reduce((current, child) => {
      const raw = child.attributes.id;
      const parsed = raw ? Number.parseInt(raw, 10) : Number.NaN;
      if (!Number.isInteger(parsed) || parsed < 256) {
        throw new OoxmlError("PPTX_PRESENTATION_INVALID", `Invalid slide id ${JSON.stringify(raw)} in ${mainPartName}`);
      }
      return Math.max(current, parsed);
    }, 255) ?? 255;

  if (maxId >= 2147483647) {
    throw new OoxmlError("PPTX_ID_EXHAUSTED", "No valid PPTX slide id remains");
  }
  return maxId + 1;
}

function appendSlideId(presentationXml: string, slideId: number, relationshipId: string): string {
  const document = parseXml(presentationXml);
  const root = document.root;
  if (root.localName !== "presentation" || root.namespaceURI !== PRESENTATION_NS) {
    throw new OoxmlError("PPTX_PRESENTATION_INVALID", "Invalid presentation root while appending slide id");
  }

  const list = root.children.find((child) => isElement(child, "sldIdLst", PRESENTATION_NS));
  const entryName = qualifiedName(list?.children.find((child) => isElement(child, "sldId", PRESENTATION_NS))?.name ?? list?.name ?? root.name, "sldId");
  const listName = qualifiedName(list?.name ?? root.name, "sldIdLst");
  const relationshipAttribute = relationshipAttributeName(root, list);
  const entry = `<${entryName} id="${slideId}" ${relationshipAttribute}="${escapeAttribute(relationshipId)}"/>`;

  if (list) {
    return appendXmlChild(presentationXml, list, entry);
  }

  const block = `<${listName}>${entry}</${listName}>`;
  const anchor = root.children.find((child) =>
    isElement(child, "sldSz", PRESENTATION_NS)
    || isElement(child, "notesSz", PRESENTATION_NS)
    || isElement(child, "defaultTextStyle", PRESENTATION_NS)
    || isElement(child, "handoutMasterIdLst", PRESENTATION_NS)
    || isElement(child, "extLst", PRESENTATION_NS)
  );

  return applyEdits(presentationXml, [{
    start: anchor?.start ?? root.closeStart,
    end: anchor?.start ?? root.closeStart,
    value: block,
  }]);
}

function appendXmlChild(xml: string, parent: XmlElement, childXml: string): string {
  if (parent.selfClosing) {
    const openTag = xml.slice(parent.start, parent.openEnd);
    if (!openTag.endsWith("/>")) {
      throw new OoxmlError("PPTX_PRESENTATION_INVALID", `Malformed self-closing element ${parent.name}`);
    }
    return applyEdits(xml, [{
      start: parent.start,
      end: parent.openEnd,
      value: `${openTag.slice(0, -2)}>${childXml}</${parent.name}>`,
    }]);
  }

  return applyEdits(xml, [{
    start: parent.closeStart,
    end: parent.closeStart,
    value: childXml,
  }]);
}

function relationshipAttributeName(root: XmlElement, list?: XmlElement): string {
  const lexical = [
    ...(list?.children ?? []),
    ...root.children,
  ]
    .map((element) => lexicalAttributeName(element, "id", OFFICE_REL_NS))
    .find((name) => name !== undefined);
  if (lexical) {
    return lexical;
  }
  const declaredPrefix = namespacePrefix(root, OFFICE_REL_NS);
  return `${declaredPrefix ?? "r"}:id`;
}

function lexicalAttributeName(element: XmlElement, localName: string, namespaceURI: string): string | undefined {
  for (const name of Object.keys(element.attributes)) {
    const local = name.includes(":") ? name.slice(name.indexOf(":") + 1) : name;
    if (local === localName && element.attributeNamespaces[name] === namespaceURI) {
      return name;
    }
  }
  return undefined;
}

function namespacePrefix(element: XmlElement, namespaceURI: string): string | undefined {
  for (const [name, value] of Object.entries(element.attributes)) {
    if (name.startsWith("xmlns:") && value === namespaceURI) {
      return name.slice("xmlns:".length);
    }
  }
  return undefined;
}

function qualifiedName(sampleName: string, localName: string): string {
  const separator = sampleName.indexOf(":");
  return separator === -1 ? localName : `${sampleName.slice(0, separator)}:${localName}`;
}

function addXmlPart(pkg: OpcPackage, partName: string, xml: string, contentType: string): void {
  addPart(pkg, partName, xml, contentType);
}

function addInternalRelationship(
  pkg: OpcPackage,
  owner: string,
  type: string,
  target: string,
): Relationship {
  return addRelationship(pkg, owner, type, target, { external: false });
}

function relativeTarget(owner: string, target: string): string {
  return posix.relative(posix.dirname(owner), target);
}

function validateAddTableArgs(rows: number, columns: number, geometry: TableGeometry): void {
  if (!Number.isSafeInteger(rows) || rows <= 0) {
    throw new OoxmlError("PPTX_ARGUMENT_INVALID", "Table rows must be a positive safe integer");
  }
  if (!Number.isSafeInteger(columns) || columns <= 0) {
    throw new OoxmlError("PPTX_ARGUMENT_INVALID", "Table columns must be a positive safe integer");
  }
  if (rows * columns > MAX_TABLE_CELLS) {
    throw new OoxmlError(
      "PPTX_ARGUMENT_INVALID",
      `Table size ${rows}x${columns} exceeds the ${MAX_TABLE_CELLS}-cell safety limit`,
    );
  }
  if (!geometry || typeof geometry !== "object") {
    throw new OoxmlError("PPTX_ARGUMENT_INVALID", "Table geometry is required");
  }

  validateGeometryInteger(geometry.x, "x", true);
  validateGeometryInteger(geometry.y, "y", true);
  validateGeometryInteger(geometry.width, "width", false);
  validateGeometryInteger(geometry.height, "height", false);
}

function validateGeometryInteger(value: number, label: string, allowZero: boolean): void {
  if (!Number.isSafeInteger(value)) {
    throw new OoxmlError("PPTX_ARGUMENT_INVALID", `Table ${label} must be a safe integer EMU value`);
  }
  if (allowZero ? value < 0 : value <= 0) {
    throw new OoxmlError(
      "PPTX_ARGUMENT_INVALID",
      `Table ${label} must be ${allowZero ? "zero or greater" : "greater than zero"}`,
    );
  }
}

function validateTextSlideArgs(title: string, subtitle: string | undefined): void {
  if (typeof title !== "string") {
    throw new OoxmlError("PPTX_ARGUMENT_INVALID", "Slide title must be a string");
  }
  if (title.length === 0) {
    throw new OoxmlError("PPTX_ARGUMENT_INVALID", "Slide title must not be empty");
  }
  if (subtitle !== undefined && typeof subtitle !== "string") {
    throw new OoxmlError("PPTX_ARGUMENT_INVALID", "Slide subtitle must be a string when provided");
  }
  if (subtitle !== undefined && subtitle.length === 0) {
    throw new OoxmlError("PPTX_ARGUMENT_INVALID", "Slide subtitle must not be empty when provided");
  }
}

function directPlaceholder(shape: XmlElement): XmlElement | undefined {
  const nvSpPr = shape.children.find((child) => isElement(child, "nvSpPr", PRESENTATION_NS));
  const nvPr = nvSpPr?.children.find((child) => isElement(child, "nvPr", PRESENTATION_NS));
  return nvPr?.children.find((child) => isElement(child, "ph", PRESENTATION_NS));
}

function requireSlideShapeTree(document: ReturnType<typeof parseXml>, partName: string): XmlElement {
  if (document.root.localName !== "sld" || document.root.namespaceURI !== PRESENTATION_NS) {
    throw new OoxmlError("PPTX_SLIDE_INVALID", `Invalid slide root in ${partName}`);
  }

  const cSld = document.root.children.find((child) => isElement(child, "cSld", PRESENTATION_NS));
  const spTree = cSld?.children.find((child) => isElement(child, "spTree", PRESENTATION_NS));
  if (!spTree) {
    throw new OoxmlError("PPTX_SLIDE_INVALID", `Slide ${partName} is missing p:cSld/p:spTree`);
  }
  return spTree;
}

function collectSlideTables(xml: string, partName: string): ResolvedSlideTable[] {
  const document = parseXml(xml);
  return collectSlideTablesFromShapeTree(requireSlideShapeTree(document, partName), partName);
}

function collectSlideTablesFromShapeTree(spTree: XmlElement, partName: string): ResolvedSlideTable[] {
  const tables: ResolvedSlideTable[] = [];

  for (const child of spTree.children) {
    if (!isElement(child, "graphicFrame", PRESENTATION_NS)) {
      continue;
    }

    const table = collectSlideTable(child, partName);
    if (table) {
      tables.push(table);
    }
  }

  return tables;
}

function collectSlideTable(frame: XmlElement, partName: string): ResolvedSlideTable | undefined {
  const graphic = frame.children.find((child) => isElement(child, "graphic", DRAWING_NS));
  const graphicData = graphic?.children.find((child) => isElement(child, "graphicData", DRAWING_NS));
  if (!graphicData) {
    return undefined;
  }
  if (graphicData.attributes.uri !== TABLE_GRAPHIC_DATA_URI) {
    return undefined;
  }

  const table = graphicData.children.find((child) => isElement(child, "tbl", DRAWING_NS));
  if (!table) {
    throw tableStructureUnsupported(partName, "graphicFrame table payload is missing a:tbl");
  }

  const grid = table.children.find((child) => isElement(child, "tblGrid", DRAWING_NS));
  const gridColumns = grid?.children.filter((child) => isElement(child, "gridCol", DRAWING_NS)) ?? [];
  const rows = table.children.filter((child) => isElement(child, "tr", DRAWING_NS));
  if (gridColumns.length === 0 || rows.length === 0) {
    throw tableStructureUnsupported(partName, "table grid must contain at least one row and one column");
  }

  const cells = rows.map((row, rowIndex) => {
    const rowCells = row.children.filter((child) => isElement(child, "tc", DRAWING_NS));
    if (rowCells.length !== gridColumns.length) {
      throw tableStructureUnsupported(
        partName,
        `table row ${rowIndex + 1} has ${rowCells.length} cells for a ${gridColumns.length}-column grid`,
      );
    }
    return rowCells;
  });

  const merged = cells.some((row) => row.some((cell) => hasMergeAttributes(cell, partName)));
  return {
    shapeId: parseGraphicFrameShapeId(frame, partName),
    table,
    gridColumns,
    rows,
    cells,
    merged,
  };
}

function parseGraphicFrameShapeId(frame: XmlElement, partName: string): number {
  const nvGraphicFramePr = frame.children.find((child) => isElement(child, "nvGraphicFramePr", PRESENTATION_NS));
  const cNvPr = nvGraphicFramePr?.children.find((child) => isElement(child, "cNvPr", PRESENTATION_NS));
  const rawId = cNvPr?.attributes.id;
  const shapeId = rawId ? Number.parseInt(rawId, 10) : Number.NaN;
  if (!Number.isInteger(shapeId) || shapeId <= 0 || shapeId > MAX_PPTX_SHAPE_ID) {
    throw new OoxmlError("PPTX_SLIDE_INVALID", `Invalid shape id ${JSON.stringify(rawId)} in ${partName}`);
  }
  return shapeId;
}

function nextSlideShapeId(document: ReturnType<typeof parseXml>, partName: string): number {
  const ids = new Set<number>();
  let maxId = 0;

  for (const cNvPr of elements(document, "cNvPr", PRESENTATION_NS)) {
    const rawId = cNvPr.attributes.id;
    const shapeId = rawId ? Number.parseInt(rawId, 10) : Number.NaN;
    if (!Number.isInteger(shapeId) || shapeId <= 0 || shapeId > MAX_PPTX_SHAPE_ID) {
      throw new OoxmlError("PPTX_SLIDE_INVALID", `Invalid shape id ${JSON.stringify(rawId)} in ${partName}`);
    }
    if (ids.has(shapeId)) {
      throw new OoxmlError("PPTX_SLIDE_INVALID", `Duplicate shape id ${shapeId} in ${partName}`);
    }
    ids.add(shapeId);
    maxId = Math.max(maxId, shapeId);
  }

  if (maxId >= MAX_PPTX_SHAPE_ID) {
    throw new OoxmlError("PPTX_ID_EXHAUSTED", `No valid shape id remains in ${partName}`);
  }
  return maxId + 1;
}

function insertShapeTreeChild(xml: string, spTree: XmlElement, childXml: string): string {
  const extLst = spTree.children.find((child) => isElement(child, "extLst", PRESENTATION_NS));
  return applyEdits(xml, [{
    start: extLst?.start ?? spTree.closeStart,
    end: extLst?.start ?? spTree.closeStart,
    value: childXml,
  }]);
}

function buildTableGraphicFrameXml(
  shapeId: number,
  rows: number,
  columns: number,
  geometry: TableGeometry,
): string {
  const name = `Table ${shapeId - 1}`;
  return [
    `<p:graphicFrame>`,
    `<p:nvGraphicFramePr>`,
    `<p:cNvPr id="${shapeId}" name="${escapeAttribute(name)}"/>`,
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>`,
    `<p:nvPr/>`,
    `</p:nvGraphicFramePr>`,
    `<p:xfrm><a:off x="${geometry.x}" y="${geometry.y}"/><a:ext cx="${geometry.width}" cy="${geometry.height}"/></p:xfrm>`,
    `<a:graphic><a:graphicData uri="${TABLE_GRAPHIC_DATA_URI}">`,
    buildTableXml(rows, columns, geometry.width, geometry.height),
    `</a:graphicData></a:graphic>`,
    `</p:graphicFrame>`,
  ].join("");
}

function buildTableXml(rows: number, columns: number, width: number, height: number): string {
  const columnWidths = distributeExtent(width, columns);
  const rowHeights = distributeExtent(height, rows);
  return [
    `<a:tbl>`,
    `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`,
    `<a:tblGrid>${columnWidths.map((columnWidth) => `<a:gridCol w="${columnWidth}"/>`).join("")}</a:tblGrid>`,
    rowHeights.map((rowHeight) => buildTableRowXml(rowHeight, columns)).join(""),
    `</a:tbl>`,
  ].join("");
}

function buildTableRowXml(height: number, columns: number): string {
  return `<a:tr h="${height}">${Array.from({ length: columns }, () => buildEmptyTableCellXml()).join("")}</a:tr>`;
}

function buildEmptyTableCellXml(): string {
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p></a:p></a:txBody><a:tcPr/></a:tc>`;
}

function distributeExtent(total: number, count: number): number[] {
  const base = Math.floor(total / count);
  return Array.from({ length: count }, (_, index) => (
    index === count - 1 ? total - (base * (count - 1)) : base
  ));
}

function assertTableIndex(index: number, limit: number, label: string): void {
  if (typeof index !== "number" || !Number.isInteger(index)) {
    throw new TypeError(`Table ${label} index must be an integer`);
  }
  if (index < 0 || index >= limit) {
    throw new RangeError(`Table ${label} index ${index} is out of range`);
  }
}

function readTableCellText(cell: XmlElement, partName: string, row: number, column: number): string {
  const paragraphs = tableCellParagraphs(cell, partName, row, column);
  const unreadableIndex = paragraphs.findIndex((paragraph) => !paragraph.readable);
  if (unreadableIndex !== -1) {
    throw cellTextUnsupported(
      partName,
      row,
      column,
      `paragraph ${unreadableIndex + 1} cannot be read faithfully`,
    );
  }
  return paragraphs.map((paragraph) => paragraph.text).join("\n");
}

function tableCellParagraphs(cell: XmlElement, partName: string, row: number, column: number): StoryParagraph[] {
  const txBody = cell.children.find((child) => isElement(child, "txBody", DRAWING_NS));
  if (!txBody) {
    throw tableStructureUnsupported(partName, `${describeTableCell(row, column)} is missing a:txBody`);
  }

  const paragraphs = txBody.children.filter((child) => isElement(child, "p", DRAWING_NS));
  const unsupported = txBody.children.find((child) => !isElement(child, "bodyPr", DRAWING_NS)
    && !isElement(child, "lstStyle", DRAWING_NS)
    && !isElement(child, "p", DRAWING_NS));
  if (unsupported || paragraphs.length === 0) {
    throw tableStructureUnsupported(partName, `${describeTableCell(row, column)} uses unsupported a:txBody topology`);
  }

  return paragraphs.map((paragraph) => analyzeParagraph(paragraph));
}

function replaceTableCellText(
  xml: string,
  cell: XmlElement,
  partName: string,
  row: number,
  column: number,
  replacement: string,
): string {
  const txBody = cell.children.find((child) => isElement(child, "txBody", DRAWING_NS));
  if (!txBody) {
    throw tableStructureUnsupported(partName, `${describeTableCell(row, column)} is missing a:txBody`);
  }

  const paragraphElements = txBody.children.filter((child) => isElement(child, "p", DRAWING_NS));
  const unsupportedTxBodyChild = txBody.children.find((child) => !isElement(child, "bodyPr", DRAWING_NS)
    && !isElement(child, "lstStyle", DRAWING_NS)
    && !isElement(child, "p", DRAWING_NS));
  if (unsupportedTxBodyChild) {
    throw tableStructureUnsupported(partName, `${describeTableCell(row, column)} uses unsupported a:txBody topology`);
  }
  if (paragraphElements.length !== 1) {
    throw cellTextUnsupported(partName, row, column, "cell must contain exactly one paragraph to edit safely");
  }

  const paragraphElement = paragraphElements[0]!;
  const paragraph = analyzeParagraph(paragraphElement);
  if (!paragraph.readable || !paragraph.replaceable) {
    throw cellTextUnsupported(partName, row, column, "cell paragraph cannot be edited safely");
  }

  const paragraphAttributes = serializeAttributes(paragraphElement.attributes);
  const pPrElement = paragraphElement.children.find((child) => isElement(child, "pPr", DRAWING_NS));
  const endParaRPrElement = paragraphElement.children.find((child) => isElement(child, "endParaRPr", DRAWING_NS));
  const pPrXml = pPrElement ? xml.slice(pPrElement.start, pPrElement.end) : "";
  const endParaRPrXml = endParaRPrElement ? xml.slice(endParaRPrElement.start, endParaRPrElement.end) : "";
  const replacementRunXml = replacement.length === 0
    ? ""
    : paragraph.editableRuns[0]
      ? buildRunXml(xml, paragraph.editableRuns[0]!, replacement)
      : buildBareDrawingRunXml(paragraphElement, replacement);

  return applyEdits(xml, [{
    start: paragraphElement.start,
    end: paragraphElement.end,
    value: `<${paragraphElement.name}${paragraphAttributes}>${pPrXml}${replacementRunXml}${endParaRPrXml}</${paragraphElement.name}>`,
  }]);
}

function buildBareDrawingRunXml(paragraphElement: XmlElement, text: string): string {
  const runName = qualifiedName(paragraphElement.name, "r");
  const textName = qualifiedName(paragraphElement.name, "t");
  const preserve = needsPreserveSpace(text) ? ' xml:space="preserve"' : "";
  return `<${runName}><${textName}${preserve}>${escapeText(text)}</${textName}></${runName}>`;
}

function describeTableCell(row: number, column: number): string {
  return `table cell (${row}, ${column})`;
}

function hasMergeAttributes(cell: XmlElement, partName: string): boolean {
  const gridSpan = parsePositiveTableInt(cell.attributes.gridSpan, "gridSpan", partName);
  const rowSpan = parsePositiveTableInt(cell.attributes.rowSpan, "rowSpan", partName);
  const hMerge = parseTableBoolean(cell.attributes.hMerge, "hMerge", partName);
  const vMerge = parseTableBoolean(cell.attributes.vMerge, "vMerge", partName);
  return gridSpan > 1 || rowSpan > 1 || hMerge || vMerge;
}

function parsePositiveTableInt(raw: string | undefined, label: string, partName: string): number {
  if (raw === undefined) {
    return 1;
  }
  if (!/^[1-9][0-9]*$/.test(raw)) {
    throw tableStructureUnsupported(partName, `Invalid ${label} value ${JSON.stringify(raw)}`);
  }
  return Number(raw);
}

function parseTableBoolean(raw: string | undefined, label: string, partName: string): boolean {
  if (raw === undefined) {
    return false;
  }
  if (raw === "1" || raw === "true") {
    return true;
  }
  if (raw === "0" || raw === "false") {
    return false;
  }
  throw tableStructureUnsupported(partName, `Invalid ${label} value ${JSON.stringify(raw)}`);
}

function tableStructureUnsupported(partName: string, detail: string): OoxmlError {
  return new OoxmlError("PPTX_TABLE_STRUCTURE_UNSUPPORTED", `PPTX table structure in ${partName} is unsupported: ${detail}`);
}

function cellTextUnsupported(partName: string, row: number, column: number, detail: string): OoxmlError {
  return new OoxmlError(
    "PPTX_UNSUPPORTED_TEXT_TOPOLOGY",
    `Table cell ${row + 1},${column + 1} in ${partName} cannot be edited safely: ${detail}`,
  );
}

function mergedTableUnsupported(partName: string, row: number, column: number): OoxmlError {
  return new OoxmlError(
    "PPTX_TABLE_MERGE_UNSUPPORTED",
    `Merged table editing is not supported for table cell ${row + 1},${column + 1} in ${partName}`,
  );
}

function staleTableHandle(partName: string, detail: string): OoxmlError {
  return new OoxmlError("PPTX_STALE_TABLE_HANDLE", `Stale PPTX table handle for ${partName}: ${detail}`);
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
  const fragments: StoryFragment[] = [];
  const editableRuns: StoryRun[] = [];
  const textParts: string[] = [];
  let offset = 0;
  let readable = true;
  let replaceable = true;

  for (const child of paragraph.children) {
    if (isElement(child, "pPr", DRAWING_NS) || isElement(child, "endParaRPr", DRAWING_NS)) {
      continue;
    }

    if (isElement(child, "r", DRAWING_NS)) {
      const textElement = child.children.find((node) => isElement(node, "t", DRAWING_NS));
      const rPrElement = child.children.find((node) => isElement(node, "rPr", DRAWING_NS));
      const supportedChildren = child.children.every(
        (node) => isElement(node, "t", DRAWING_NS) || isElement(node, "rPr", DRAWING_NS),
      );

      if (!textElement || !supportedChildren || child.children.filter((node) => isElement(node, "t", DRAWING_NS)).length !== 1) {
        readable = false;
        replaceable = false;
        continue;
      }

      const text = textElement.text;
      const attrs = { ...(rPrElement?.attributes ?? {}) };
      fragments.push({ text, attrs });
      editableRuns.push({
        element: child,
        textElement,
        rPrElement,
        text,
        start: offset,
        end: offset + text.length,
        attrs,
      });
      textParts.push(text);
      offset += text.length;
      continue;
    }

    if (isElement(child, "br", DRAWING_NS)) {
      const rPrElement = child.children.find((node) => isElement(node, "rPr", DRAWING_NS));
      const supportedChildren = child.children.every((node) => isElement(node, "rPr", DRAWING_NS));
      if (!supportedChildren) {
        readable = false;
      }

      const text = "\n";
      fragments.push({ text, attrs: { ...(rPrElement?.attributes ?? {}) } });
      textParts.push(text);
      offset += text.length;
      replaceable = false;
      continue;
    }

    if (isElement(child, "fld", DRAWING_NS)) {
      const textElement = child.children.find((node) => isElement(node, "t", DRAWING_NS));
      const rPrElement = child.children.find((node) => isElement(node, "rPr", DRAWING_NS));
      const supportedChildren = child.children.every(
        (node) => isElement(node, "t", DRAWING_NS)
          || isElement(node, "rPr", DRAWING_NS)
          || isElement(node, "endParaRPr", DRAWING_NS),
      );
      if (!textElement || !supportedChildren || child.children.filter((node) => isElement(node, "t", DRAWING_NS)).length !== 1) {
        readable = false;
        replaceable = false;
        continue;
      }

      const text = textElement.text;
      fragments.push({ text, attrs: { ...(rPrElement?.attributes ?? {}) } });
      textParts.push(text);
      offset += text.length;
      replaceable = false;
      continue;
    }

    readable = false;
    replaceable = false;
  }

  return {
    element: paragraph,
    text: textParts.join(""),
    fragments,
    editableRuns,
    readable,
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
  const touched = paragraph.editableRuns.filter((run) => run.end > matchStart && run.start < matchEnd);
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

function assertReadableParagraphs(partName: string, paragraphs: StoryParagraph[]): void {
  const unreadableIndex = paragraphs.findIndex((paragraph) => !paragraph.readable);
  if (unreadableIndex !== -1) {
    throw new OoxmlError(
      "PPTX_UNSUPPORTED_TEXT_TOPOLOGY",
      `Paragraph ${unreadableIndex + 1} in ${partName} cannot be read faithfully`,
    );
  }
}

function staleAnchor(partName: string, detail: string): OoxmlError {
  return new OoxmlError("PPTX_STALE_ANCHOR", `Stale PPTX text anchor for ${partName}: ${detail}`);
}
