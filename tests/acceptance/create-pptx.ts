import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { Presentation } from "../../src/pptx/index.ts";
import { findPlaceholderText } from "../../src/pptx/placeholders.ts";
import { parseXml } from "../../src/xml/index.ts";
import { fixturePath, F } from "../../scripts/fixture-inputs.ts";

const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

const OFFICE_DOCUMENT_RELATIONSHIP = `${OFFICE_REL_NS}/officeDocument`;
const SLIDE_RELATIONSHIP = `${OFFICE_REL_NS}/slide`;
const SLIDE_MASTER_RELATIONSHIP = `${OFFICE_REL_NS}/slideMaster`;
const SLIDE_LAYOUT_RELATIONSHIP = `${OFFICE_REL_NS}/slideLayout`;
const THEME_RELATIONSHIP = `${OFFICE_REL_NS}/theme`;
const VIEW_PROPS_RELATIONSHIP = `${OFFICE_REL_NS}/viewProps`;
const PRES_PROPS_RELATIONSHIP = `${OFFICE_REL_NS}/presProps`;
const TABLE_STYLES_RELATIONSHIP = `${OFFICE_REL_NS}/tableStyles`;

const SHARED_TITLE_FIXTURE = fixturePath(F.shared.titleAndSubtitle);
const encoder = new TextEncoder();
const decoder = new TextDecoder();

const CONTENT_TYPES_PART = "[Content_Types].xml";
const ROOT_RELS_PART = "_rels/.rels";
const PRESENTATION_PART = "ppt/presentation.xml";
const PRESENTATION_RELS_PART = "ppt/_rels/presentation.xml.rels";
const PRES_PROPS_PART = "ppt/presProps.xml";
const VIEW_PROPS_PART = "ppt/viewProps.xml";
const TABLE_STYLES_PART = "ppt/tableStyles.xml";
const THEME_PART = "ppt/theme/theme1.xml";
const SLIDE_MASTER_PART = "ppt/slideMasters/slideMaster1.xml";
const SLIDE_MASTER_RELS_PART = "ppt/slideMasters/_rels/slideMaster1.xml.rels";
const SLIDE_LAYOUT_PART = "ppt/slideLayouts/slideLayout1.xml";
const SLIDE_LAYOUT_RELS_PART = "ppt/slideLayouts/_rels/slideLayout1.xml.rels";
const SLIDE1_PART = "ppt/slides/slide1.xml";
const SLIDE1_RELS_PART = "ppt/slides/_rels/slide1.xml.rels";
const SLIDE2_PART = "ppt/slides/slide2.xml";
const SLIDE2_RELS_PART = "ppt/slides/_rels/slide2.xml.rels";

export type MinimalCreateResult = {
  savedBytes: Uint8Array;
  savedPartNames: string[];
  reopenedTexts: string[][];
  slideCount: number;
  slideMasterCount: number;
  slideLayoutCount: number;
  mainPartName: string;
  rootOfficeDocumentTarget?: string;
  presentationRelationshipTypes: string[];
  presentationSlideTargets: string[];
  presentationSlideMasterTarget?: string;
  presentationThemeTarget?: string;
  presentationPresPropsTarget?: string;
  presentationViewPropsTarget?: string;
  presentationTableStylesTarget?: string;
  slideLayoutTargets: string[];
  masterRelationshipTypes: string[];
  masterLayoutTargets: string[];
  masterThemeTarget?: string;
  layoutRelationshipTypes: string[];
  layoutMasterTarget?: string;
  contentTypeOverrides: string[];
  reopenedDiff: { added: string[]; changed: string[]; removed: string[] };
};

export type AppendPreservationResult = {
  sourceBytes: Uint8Array;
  bytesAfterAppend: Uint8Array;
  savedBytes: Uint8Array;
  savedPartNames: string[];
  reopenedTexts: string[][];
  originalSlide1Xml: string;
  slide1XmlAfterAppend: string;
  expectedSlide1Xml: string;
  finalSlide1Xml: string;
  originalSlide1Rels: Uint8Array;
  finalSlide1Rels: Uint8Array;
  originalTitleAnchorVersion: number;
  originalTitleAnchorParagraphIndex: number;
  afterAppendTitleAnchorVersion: number;
  afterAppendTitleAnchorParagraphIndex: number;
  presentationSlideTargets: string[];
  slideLayoutTargets: string[];
  diff: { added: string[]; changed: string[]; removed: string[] };
};

export type CreateRefusalResult = {
  invalidBefore: Uint8Array;
  invalidAfter: Uint8Array;
  invalidDiff: { added: string[]; changed: string[]; removed: string[] };
  invalidError: OoxmlError;
  unsafeBefore: Uint8Array;
  unsafeAfter: Uint8Array;
  unsafeDiff: { added: string[]; changed: string[]; removed: string[] };
  unsafeError: OoxmlError;
};

type AcceptanceState = {
  minimal?: MinimalCreateResult;
  append?: AppendPreservationResult;
  refusal?: CreateRefusalResult;
};

export async function runCreateMinimalScenario(): Promise<MinimalCreateResult> {
  const presentation = Presentation.create();
  presentation.addTextSlide("Created title 1", "Created subtitle 1");
  presentation.addTextSlide("Created title 2", "Created subtitle 2");

  const savedBytes = await saveToTempAndReadBytes(presentation);
  const savedParts = readZip(savedBytes);
  const reopened = await Presentation.open(savedBytes);
  const packageRels = reopened.package.relationships(PRESENTATION_PART);
  const masterRels = reopened.package.relationships(SLIDE_MASTER_PART);
  const layoutRels = reopened.package.relationships(SLIDE_LAYOUT_PART);

  return {
    savedBytes,
    savedPartNames: [...savedParts.keys()].sort(),
    reopenedTexts: collectSlidePlaceholderTexts(reopened),
    slideCount: reopened.slideCount,
    slideMasterCount: reopened.slideMasterCount,
    slideLayoutCount: reopened.slideLayoutCount,
    mainPartName: reopened.package.mainPart(),
    rootOfficeDocumentTarget: reopened.package.related("", "officeDocument"),
    presentationRelationshipTypes: packageRels.map((relationship) => relationship.type),
    presentationSlideTargets: packageRels
      .filter((relationship) => relationship.type === SLIDE_RELATIONSHIP)
      .map((relationship) => required(relationship.resolved, "presentation slide relationship target")),
    presentationSlideMasterTarget: reopened.package.related(PRESENTATION_PART, "slideMaster"),
    presentationThemeTarget: reopened.package.related(PRESENTATION_PART, "theme"),
    presentationPresPropsTarget: reopened.package.related(PRESENTATION_PART, "presProps"),
    presentationViewPropsTarget: reopened.package.related(PRESENTATION_PART, "viewProps"),
    presentationTableStylesTarget: reopened.package.related(PRESENTATION_PART, "tableStyles"),
    slideLayoutTargets: reopened.slides.map((slide) => required(
      reopened.package.related(slide.partName, "slideLayout"),
      `slide layout relationship for ${slide.partName}`,
    )),
    masterRelationshipTypes: masterRels.map((relationship) => relationship.type),
    masterLayoutTargets: masterRels
      .filter((relationship) => relationship.type === SLIDE_LAYOUT_RELATIONSHIP)
      .map((relationship) => required(relationship.resolved, "master layout relationship target")),
    masterThemeTarget: reopened.package.related(SLIDE_MASTER_PART, "theme"),
    layoutRelationshipTypes: layoutRels.map((relationship) => relationship.type),
    layoutMasterTarget: reopened.package.related(SLIDE_LAYOUT_PART, "slideMaster"),
    contentTypeOverrides: readContentTypeOverrides(savedParts),
    reopenedDiff: reopened.package.diff(),
  };
}

export async function runAppendPreservationScenario(
  source: Uint8Array = readFixtureBytesSync(SHARED_TITLE_FIXTURE),
): Promise<AppendPreservationResult> {
  const sourceParts = readZip(source);
  const originalSlide1Xml = decodeRequiredPart(sourceParts, SLIDE1_PART);
  const originalSlide1Rels = required(sourceParts.get(SLIDE1_RELS_PART), `Missing ${SLIDE1_RELS_PART}`);

  const presentation = await Presentation.open(source);
  const slide = required(presentation.slides[0], "existing title slide");
  const title = findPlaceholderText(slide, "title");

  presentation.addTextSlide("Appended title", "Appended subtitle");

  const afterAppendTitle = findPlaceholderText(slide, "title");
  const bytesAfterAppend = presentation.package.toBytes();
  const partsAfterAppend = readZip(bytesAfterAppend);
  const slide1XmlAfterAppend = decodeRequiredPart(partsAfterAppend, SLIDE1_PART);

  slide.replaceTextAt(title.anchor, title.text, "Edited original title");

  const savedBytes = await saveToTempAndReadBytes(presentation);
  const savedParts = readZip(savedBytes);
  const reopened = await Presentation.open(savedBytes);

  return {
    sourceBytes: source,
    bytesAfterAppend,
    savedBytes,
    savedPartNames: [...savedParts.keys()].sort(),
    reopenedTexts: collectSlidePlaceholderTexts(reopened),
    originalSlide1Xml,
    slide1XmlAfterAppend,
    expectedSlide1Xml: originalSlide1Xml.replace(title.text, "Edited original title"),
    finalSlide1Xml: decodeRequiredPart(savedParts, SLIDE1_PART),
    originalSlide1Rels,
    finalSlide1Rels: required(savedParts.get(SLIDE1_RELS_PART), `Missing ${SLIDE1_RELS_PART}`),
    originalTitleAnchorVersion: title.anchor.version,
    originalTitleAnchorParagraphIndex: title.anchor.paragraphIndex,
    afterAppendTitleAnchorVersion: afterAppendTitle.anchor.version,
    afterAppendTitleAnchorParagraphIndex: afterAppendTitle.anchor.paragraphIndex,
    presentationSlideTargets: reopened.package.relationships(PRESENTATION_PART)
      .filter((relationship) => relationship.type === SLIDE_RELATIONSHIP)
      .map((relationship) => required(relationship.resolved, "presentation slide relationship target")),
    slideLayoutTargets: reopened.slides.map((currentSlide) => required(
      reopened.package.related(currentSlide.partName, "slideLayout"),
      `slide layout relationship for ${currentSlide.partName}`,
    )),
    diff: presentation.package.diff(),
  };
}

export function createUnsafeTitleLayoutFixture(): Uint8Array {
  const parts = readZip(readFixtureBytesSync(SHARED_TITLE_FIXTURE));
  const layoutXml = decodeRequiredPart(parts, SLIDE_LAYOUT_PART);
  const unsafeLayoutXml = layoutXml.replace('<p:ph type="subTitle" idx="1"/>', '<p:ph type="body" idx="1"/>');
  if (unsafeLayoutXml === layoutXml) {
    throw new Error("expected title layout fixture to expose a direct subtitle placeholder");
  }
  parts.set(SLIDE_LAYOUT_PART, encoder.encode(unsafeLayoutXml));
  return writeZip(parts);
}

export async function runCreateRefusalScenario(
  unsafeSource: Uint8Array = createUnsafeTitleLayoutFixture(),
): Promise<CreateRefusalResult> {
  const invalidPresentation = Presentation.create();
  const invalidBefore = invalidPresentation.package.toBytes();
  const invalidError = captureError(() => {
    (invalidPresentation.addTextSlide as unknown as (title: unknown, subtitle?: unknown) => unknown)(7, "Bad subtitle");
  });

  const unsafeBefore = unsafeSource.slice();
  const unsafePresentation = await Presentation.open(unsafeBefore);
  const unsafeError = captureError(() => {
    unsafePresentation.addTextSlide("Unsafe next title", "Unsafe next subtitle");
  });

  return {
    invalidBefore,
    invalidAfter: invalidPresentation.package.toBytes(),
    invalidDiff: invalidPresentation.package.diff(),
    invalidError,
    unsafeBefore,
    unsafeAfter: unsafePresentation.package.toBytes(),
    unsafeDiff: unsafePresentation.package.diff(),
    unsafeError,
  };
}

export const bindings: StepBinding[] = [
  {
    pattern: /^a new native PPTX presentation is created from scratch$/,
    run: (context) => {
      scenarioState(context).minimal = undefined;
    },
  },
  {
    pattern: /^PPTX adds two text slides and saves then reopens the package$/,
    run: async (context) => {
      scenarioState(context).minimal = await runCreateMinimalScenario();
    },
  },
  {
    pattern: /^PPTX reopens both slides in order with title and subtitle placeholder text and valid minimal relationships$/,
    run: (context) => {
      const result = required(scenarioState(context).minimal, "minimal create result");
      assertMinimalCreateResult(result);
    },
  },
  {
    pattern: /^a compatible existing PPTX title-slide deck is prepared$/,
    run: (context) => {
      scenarioState(context).append = undefined;
    },
  },
  {
    pattern: /^PPTX appends a new text slide and edits the original title through its earlier anchor$/,
    run: async (context) => {
      scenarioState(context).append = await runAppendPreservationScenario();
    },
  },
  {
    pattern: /^PPTX keeps the original slide content ordered and preserved except for the anchored edit and gives the new slide independent relationships$/,
    run: (context) => {
      const result = required(scenarioState(context).append, "append result");
      assertAppendPreservationResult(result);
    },
  },
  {
    pattern: /^invalid-argument and unsafe-layout PPTX fixtures are prepared$/,
    run: (context) => {
      scenarioState(context).refusal = undefined;
    },
  },
  {
    pattern: /^PPTX attempts unsupported creation edits on those fixtures$/,
    run: async (context) => {
      scenarioState(context).refusal = await runCreateRefusalScenario();
    },
  },
  {
    pattern: /^PPTX refuses both requests with rollback and leaves their package bytes unchanged$/,
    run: (context) => {
      const result = required(scenarioState(context).refusal, "refusal result");
      assertCreateRefusalResult(result);
    },
  },
];

function scenarioState(context: Record<string, unknown>): AcceptanceState {
  return context.state as AcceptanceState;
}

function assertMinimalCreateResult(result: MinimalCreateResult): void {
  assert.equal(result.slideCount, 2, "created deck should reopen two slides");
  assert.equal(result.slideMasterCount, 1, "created deck should expose one slide master");
  assert.equal(result.slideLayoutCount, 1, "created deck should expose one slide layout");
  assert.equal(result.mainPartName, PRESENTATION_PART, "root officeDocument link should resolve to the presentation part");
  assert.equal(result.rootOfficeDocumentTarget, PRESENTATION_PART, "root officeDocument relationship should target ppt/presentation.xml");
  assert.deepEqual(result.reopenedTexts, [
    ["Created title 1", "Created subtitle 1"],
    ["Created title 2", "Created subtitle 2"],
  ], "created slides should round-trip ordered title/subtitle placeholder text");
  assert.deepEqual(result.savedPartNames, minimalTwoSlidePartNames(), "created deck should contain only the minimal authored parts");
  assert.deepEqual(result.contentTypeOverrides, [
    "/ppt/presProps.xml",
    "/ppt/presentation.xml",
    "/ppt/slideLayouts/slideLayout1.xml",
    "/ppt/slideMasters/slideMaster1.xml",
    "/ppt/slides/slide1.xml",
    "/ppt/slides/slide2.xml",
    "/ppt/tableStyles.xml",
    "/ppt/theme/theme1.xml",
    "/ppt/viewProps.xml",
  ], "content types should enumerate the minimal authored presentation, theme, master, layout and slide parts");
  assert.deepEqual(result.presentationRelationshipTypes.sort(), [
    PRES_PROPS_RELATIONSHIP,
    SLIDE_MASTER_RELATIONSHIP,
    SLIDE_RELATIONSHIP,
    SLIDE_RELATIONSHIP,
    TABLE_STYLES_RELATIONSHIP,
    THEME_RELATIONSHIP,
    VIEW_PROPS_RELATIONSHIP,
  ].sort(), "presentation relationships should link only the minimal slide/master/theme support graph");
  assert.equal(result.presentationSlideMasterTarget, SLIDE_MASTER_PART, "presentation should link directly to slideMaster1.xml");
  assert.equal(result.presentationThemeTarget, THEME_PART, "presentation should link directly to theme1.xml");
  assert.equal(result.presentationPresPropsTarget, PRES_PROPS_PART, "presentation should link directly to presProps.xml");
  assert.equal(result.presentationViewPropsTarget, VIEW_PROPS_PART, "presentation should link directly to viewProps.xml");
  assert.equal(result.presentationTableStylesTarget, TABLE_STYLES_PART, "presentation should link directly to tableStyles.xml");
  assert.deepEqual(result.presentationSlideTargets, [SLIDE1_PART, SLIDE2_PART], "presentation slide order should match authored relationship order");
  assert.deepEqual(result.slideLayoutTargets, [SLIDE_LAYOUT_PART, SLIDE_LAYOUT_PART], "each slide should link to the owned title layout");
  assert.deepEqual(result.masterRelationshipTypes.sort(), [
    SLIDE_LAYOUT_RELATIONSHIP,
    THEME_RELATIONSHIP,
  ].sort(), "slide master should link only to its layout and theme");
  assert.deepEqual(result.masterLayoutTargets, [SLIDE_LAYOUT_PART], "slide master should link to slideLayout1.xml");
  assert.equal(result.masterThemeTarget, THEME_PART, "slide master should link to theme1.xml");
  assert.deepEqual(result.layoutRelationshipTypes, [SLIDE_MASTER_RELATIONSHIP], "title layout should link back to its slide master only");
  assert.equal(result.layoutMasterTarget, SLIDE_MASTER_PART, "title layout should link back to slideMaster1.xml");
  assert.deepEqual(result.reopenedDiff, { added: [], changed: [], removed: [] }, "reopening the saved deck should be a no-op package read");
}

function assertAppendPreservationResult(result: AppendPreservationResult): void {
  assert.deepEqual(result.reopenedTexts, [
    ["Edited original title", "Original subtitle"],
    ["Appended title", "Appended subtitle"],
  ], "append flow should preserve slide order while applying only the anchored first-slide edit");
  assert.equal(result.originalTitleAnchorVersion, result.afterAppendTitleAnchorVersion, "appending a new slide should preserve earlier first-slide anchors");
  assert.equal(result.originalTitleAnchorParagraphIndex, result.afterAppendTitleAnchorParagraphIndex, "appending a new slide should not move the original title paragraph anchor");
  assert.equal(result.slide1XmlAfterAppend, result.originalSlide1Xml, "appending a later slide must leave the original slide XML byte-exact before editing");
  assert.equal(result.finalSlide1Xml, result.expectedSlide1Xml, "editing through the earlier anchor should change only the first slide title bytes");
  assertBytesEqual(result.finalSlide1Rels, result.originalSlide1Rels, "editing through the earlier anchor must not disturb the original slide relationship bytes");
  assert.ok(result.savedPartNames.includes(SLIDE1_RELS_PART), "saved deck should retain the original slide relationships part");
  assert.ok(result.savedPartNames.includes(SLIDE2_RELS_PART), "saved deck should add an independent relationships part for the appended slide");
  assert.deepEqual(result.presentationSlideTargets, [SLIDE1_PART, SLIDE2_PART], "presentation slide order should keep the original slide first and append the new slide second");
  assert.deepEqual(result.slideLayoutTargets, [SLIDE_LAYOUT_PART, SLIDE_LAYOUT_PART], "both original and appended slides should resolve their own layout link to the same safe title layout");
  assert.deepEqual(result.diff, {
    added: [SLIDE2_RELS_PART, SLIDE2_PART],
    changed: [CONTENT_TYPES_PART, PRESENTATION_RELS_PART, PRESENTATION_PART, SLIDE1_PART],
    removed: [],
  }, "append and anchored edit should touch only presentation scaffolding, the new slide and the edited first slide XML");
}

function assertCreateRefusalResult(result: CreateRefusalResult): void {
  assert.equal(result.invalidError.code, "PPTX_ARGUMENT_INVALID", "invalid create arguments should refuse with a stable code");
  assertBytesEqual(result.invalidAfter, result.invalidBefore, "invalid create arguments must not mutate package bytes");
  assert.deepEqual(result.invalidDiff, { added: [], changed: [], removed: [] }, "invalid create arguments must leave package diff empty");
  assert.equal(result.unsafeError.code, "PPTX_LAYOUT_UNSAFE", "unsafe existing layouts should refuse with a stable code");
  assertBytesEqual(result.unsafeAfter, result.unsafeBefore, "unsafe layout refusal must roll back package bytes");
  assert.deepEqual(result.unsafeDiff, { added: [], changed: [], removed: [] }, "unsafe layout refusal must leave package diff empty");
}

function minimalTwoSlidePartNames(): string[] {
  return [
    CONTENT_TYPES_PART,
    ROOT_RELS_PART,
    PRESENTATION_RELS_PART,
    PRES_PROPS_PART,
    PRESENTATION_PART,
    SLIDE_LAYOUT_RELS_PART,
    SLIDE_LAYOUT_PART,
    SLIDE_MASTER_RELS_PART,
    SLIDE_MASTER_PART,
    SLIDE1_RELS_PART,
    SLIDE2_RELS_PART,
    SLIDE1_PART,
    SLIDE2_PART,
    TABLE_STYLES_PART,
    THEME_PART,
    VIEW_PROPS_PART,
  ].sort();
}

function collectSlidePlaceholderTexts(presentation: Presentation): string[][] {
  return presentation.slides.map((slide) => [
    findPlaceholderText(slide, "title").text,
    findPlaceholderText(slide, "subtitle").text,
  ]);
}

function readContentTypeOverrides(parts: ReadonlyMap<string, Uint8Array>): string[] {
  const document = parseXml(decodeRequiredPart(parts, CONTENT_TYPES_PART));
  if (document.root.localName !== "Types" || document.root.namespaceURI !== CONTENT_TYPES_NS) {
    throw new Error(`invalid content types root in ${CONTENT_TYPES_PART}`);
  }
  return document.root.children
    .filter((child) => child.localName === "Override" && child.namespaceURI === CONTENT_TYPES_NS)
    .map((child) => required(child.attributes.PartName, "content type override name"))
    .sort();
}

async function saveToTempAndReadBytes(presentation: Presentation): Promise<Uint8Array> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-pptx-create-"));
  try {
    const path = join(root, "created.pptx");
    await presentation.save(path);
    return Uint8Array.from(await Bun.file(path).bytes());
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}

function captureError(operation: () => unknown): OoxmlError {
  try {
    operation();
  } catch (error) {
    if (error instanceof OoxmlError) {
      return error;
    }
    throw error;
  }
  throw new Error("expected OoxmlError");
}

function decodeRequiredPart(parts: ReadonlyMap<string, Uint8Array>, partName: string): string {
  return decoder.decode(required(parts.get(partName), `Missing ${partName}`));
}

function readFixtureBytesSync(path: string): Uint8Array {
  return Uint8Array.from(readFileSync(path));
}

function assertBytesEqual(actual: Uint8Array, expected: Uint8Array, message: string): void {
  if (actual.length !== expected.length) {
    throw new Error(`${message}\nexpected byte length ${expected.length}, got ${actual.length}`);
  }
  for (let index = 0; index < actual.length; index += 1) {
    if (actual[index] !== expected[index]) {
      throw new Error(`${message}\nfirst byte mismatch at index ${index}`);
    }
  }
}

function required<T>(value: T | undefined, message: string): T {
  assert.ok(value !== undefined, message);
  return value;
}
