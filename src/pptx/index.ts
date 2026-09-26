import { posix } from "node:path";

import { OoxmlError } from "../errors.ts";
import { addPart, addRelationship, nextPartName } from "../opc/graph.ts";
import { OpcPackage, type Relationship } from "../opc/package.ts";
import { attribute, applyEdits, elements, escapeAttribute, escapeText, parseXml, type XmlElement } from "../xml/index.ts";

const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
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
    `<a:tblStyleLst xmlns:a="${DRAWING_NS}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`,
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
