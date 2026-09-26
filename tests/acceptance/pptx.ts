import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { Presentation } from "../../src/pptx/index.ts";
import { parseXml } from "../../src/xml/index.ts";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const GO_NOTES_FIXTURE = join(
  import.meta.dir,
  "../../fixtures/go-ooxml/testdata/pptx/notes.pptx",
);
const GO_MINIMAL_FIXTURE = join(
  import.meta.dir,
  "../../fixtures/go-ooxml/testdata/pptx/minimal.pptx",
);
const PYTHON_TITLE_FIXTURE = join(
  import.meta.dir,
  "../../fixtures/python-office-mcp-server/tests/_templates/testdata/pptx/title_slide.pptx",
);

const FRAGMENTED_TITLE_PARAGRAPH = [
  "<a:p>",
  '<a:r><a:rPr b="1"/><a:t>Fran</a:t></a:r>',
  '<a:r><a:rPr i="1"/><a:t>ken</a:t></a:r>',
  '<a:r><a:rPr u="sng"/><a:t>stein</a:t></a:r>',
  "</a:p>",
].join("");

const BREAK_FIELD_TITLE_PARAGRAPH = [
  "<a:p>",
  '<a:r><a:rPr b="1"/><a:t>Chapter</a:t></a:r>',
  '<a:br><a:rPr lang="en-US"/></a:br>',
  '<a:fld id="{00000000-0000-0000-0000-000000000007}" type="slidenum"><a:rPr i="1"/><a:t>7</a:t></a:fld>',
  '<a:r><a:rPr u="sng"/><a:t> Notes</a:t></a:r>',
  "</a:p>",
].join("");

const COMPLEX_NOTES_PARAGRAPHS = [
  '<a:p><a:r><a:t>Key themes:</a:t></a:r><a:br/><a:r><a:t>creation, responsibility, isolation</a:t></a:r></a:p>',
  "<a:p></a:p>",
  '<a:p><a:r><a:t>Visible date: </a:t></a:r><a:fld id="{00000000-0000-0000-0000-000000000099}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>2026-03-12</a:t></a:fld></a:p>',
].join("");

type OrderedNotesFixture = {
  orderedBytes: Uint8Array;
  minimalSource: Uint8Array;
};

type OrderedNotesResult = {
  titles: string[];
  notes: string[];
  orderedBytes: Uint8Array;
  orderedBytesAfterRead: Uint8Array;
  orderedDiff: { added: string[]; changed: string[]; removed: string[] };
  minimalSource: Uint8Array;
  minimalError: OoxmlError;
  minimalBytesAfterRefusal: Uint8Array;
  minimalNames: string[];
};

type ReadableUnsupportedResult = {
  sourceBytes: Uint8Array;
  paragraphText: string;
  paragraphRuns: Array<{ text: string; attrs: Record<string, string> }>;
  error: OoxmlError;
  bytesAfterRefusal: Uint8Array;
  diffAfterRefusal: { added: string[]; changed: string[]; removed: string[] };
};

type CrossRunResult = {
  sourceBytes: Uint8Array;
  savedBytes: Uint8Array;
  reopenedTitle: string;
  changedParts: { added: string[]; changed: string[]; removed: string[] };
  customMemberBytes: Uint8Array;
  docPropsBytes: Uint8Array;
  titleRuns: Array<{ text: string; attrs: Record<string, string> }>;
};

type StaleAnchorResult = {
  error: OoxmlError;
  bytesAfterSuccess: Uint8Array;
  bytesAfterRefusal: Uint8Array;
  reopenedTitle: string;
};

type AcceptanceState = {
  orderedNotesFixture?: OrderedNotesFixture;
  orderedNotes?: OrderedNotesResult;
  readableUnsupportedFixture?: Uint8Array;
  readableUnsupported?: ReadableUnsupportedResult;
  crossRunFixture?: Uint8Array;
  crossRun?: CrossRunResult;
  staleAnchorFixture?: Uint8Array;
  staleAnchor?: StaleAnchorResult;
};

export function prepareOrderedNotesFixture(): OrderedNotesFixture {
  return {
    orderedBytes: buildRelationshipOrderedNotesFixture(),
    minimalSource: Uint8Array.from(readFixtureBytesSync(GO_MINIMAL_FIXTURE)),
  };
}

export async function runOrderedNotesScenario(
  fixture: OrderedNotesFixture = prepareOrderedNotesFixture(),
): Promise<OrderedNotesResult> {
  const orderedPresentation = await Presentation.open(fixture.orderedBytes);
  const titles = orderedPresentation.slides.map((slide) => {
    const paragraphs = slide.inspectText("acceptance.order.notes");
    return paragraphs[0]?.text ?? "";
  });
  const notes = [
    orderedPresentation.slides[0]?.readNotesText() ?? "",
    orderedPresentation.slides.at(-1)?.readNotesText() ?? "",
  ];

  const minimalPresentation = await Presentation.open(fixture.minimalSource);
  let minimalError: OoxmlError | undefined;
  try {
    minimalPresentation.slides[0]?.readNotesText();
  } catch (error) {
    if (error instanceof OoxmlError) {
      minimalError = error;
    } else {
      throw error;
    }
  }
  if (!minimalError) {
    throw new Error("expected missing notes refusal");
  }

  return {
    titles,
    notes,
    orderedBytes: fixture.orderedBytes,
    orderedBytesAfterRead: orderedPresentation.package.toBytes(),
    orderedDiff: orderedPresentation.package.diff(),
    minimalSource: fixture.minimalSource,
    minimalError,
    minimalBytesAfterRefusal: minimalPresentation.package.toBytes(),
    minimalNames: minimalPresentation.package.names(),
  };
}

export function prepareReadableUnsupportedFixture(): Uint8Array {
  return buildReadableUnsupportedSlideFixture();
}

export async function runReadableUnsupportedScenario(
  source: Uint8Array = prepareReadableUnsupportedFixture(),
): Promise<ReadableUnsupportedResult> {
  const presentation = await Presentation.open(source);
  const slide = required(presentation.slides[0], "readable unsupported fixture should contain slide 1");
  const paragraph = required(
    slide.inspectText("acceptance.readable.unsupported")[0],
    "expected discoverable title paragraph",
  );

  let refusal: OoxmlError | undefined;
  try {
    slide.replaceTextAt(paragraph.anchor, "Chapter", "Section");
  } catch (error) {
    if (error instanceof OoxmlError) {
      refusal = error;
    } else {
      throw error;
    }
  }
  if (!refusal) {
    throw new Error("expected unsupported-topology refusal");
  }

  return {
    sourceBytes: source,
    paragraphText: paragraph.text,
    paragraphRuns: paragraph.runs,
    error: refusal,
    bytesAfterRefusal: presentation.package.toBytes(),
    diffAfterRefusal: presentation.package.diff(),
  };
}

export function prepareCrossRunFixture(): Uint8Array {
  return buildFragmentedTitleFixture();
}

export async function runCrossRunReplacementScenario(
  source: Uint8Array = prepareCrossRunFixture(),
): Promise<CrossRunResult> {
  const presentation = await Presentation.open(source);
  const slide = required(presentation.slides[0], "fragmented fixture should contain slide 1");
  const paragraph = slide.inspectText("acceptance.cross.run").find((item) => item.text === "Frankenstein");
  if (!paragraph) {
    throw new Error("expected title paragraph to be discoverable");
  }

  slide.replaceTextAt(paragraph.anchor, "anken", "iend");
  const changedParts = presentation.package.diff();
  const savedBytes = await saveToTempAndReadBytes(presentation);
  const reopened = await Presentation.open(savedBytes);
  const reopenedTitle = reopened.slides[0]?.inspectText("acceptance.cross.run.reopen")[0]?.text ?? "";
  const savedParts = readZip(savedBytes);
  const customMemberBytes = savedParts.get("custom/data.bin");
  const docPropsBytes = savedParts.get("docProps/app.xml");
  if (!customMemberBytes || !docPropsBytes) {
    throw new Error("saved package is missing expected untouched members");
  }

  return {
    sourceBytes: source,
    savedBytes,
    reopenedTitle,
    changedParts,
    customMemberBytes,
    docPropsBytes,
    titleRuns: readFirstParagraphRuns(
      decoder.decode(savedParts.get("ppt/slides/slide1.xml") ?? new Uint8Array()),
    ),
  };
}

export function prepareStaleAnchorFixture(): Uint8Array {
  return Uint8Array.from(readFixtureBytesSync(PYTHON_TITLE_FIXTURE));
}

export async function runStaleAnchorScenario(
  source: Uint8Array = prepareStaleAnchorFixture(),
): Promise<StaleAnchorResult> {
  const presentation = await Presentation.open(source);
  const slide = required(presentation.slides[0], "stale-anchor fixture should contain slide 1");
  const paragraph = slide.inspectText("acceptance.stale.anchor").find((item) => item.text === "Frankenstein");
  if (!paragraph) {
    throw new Error("expected discoverable title paragraph for stale-anchor test");
  }

  slide.replaceTextAt(paragraph.anchor, "Frankenstein", "Creature");
  const bytesAfterSuccess = presentation.package.toBytes();

  let refusal: OoxmlError | undefined;
  try {
    slide.replaceTextAt(paragraph.anchor, "Frankenstein", "Monster");
  } catch (error) {
    if (error instanceof OoxmlError) {
      refusal = error;
    } else {
      throw error;
    }
  }
  if (!refusal) {
    throw new Error("expected stale anchor refusal");
  }

  const bytesAfterRefusal = presentation.package.toBytes();
  const reopened = await Presentation.open(bytesAfterRefusal);

  return {
    error: refusal,
    bytesAfterSuccess,
    bytesAfterRefusal,
    reopenedTitle: reopened.slides[0]?.inspectText("acceptance.stale.anchor.reopen")[0]?.text ?? "",
  };
}

export const bindings: StepBinding[] = [
  {
    pattern: /^PPTX ordered notes fixtures are prepared$/,
    run: (context) => {
      const state = scenarioState(context);
      state.orderedNotesFixture = prepareOrderedNotesFixture();
      state.orderedNotes = undefined;
    },
  },
  {
    pattern: /^PPTX opens the reordered notes fixture and probes notes reads$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.orderedNotes = await runOrderedNotesScenario(required(state.orderedNotesFixture, "orderedNotesFixture"));
    },
  },
  {
    pattern: /^PPTX keeps slide order, notes blank lines, and notes reads non-mutating without creating missing notes parts$/,
    run: (context) => {
      const state = scenarioState(context);
      const result = required(state.orderedNotes, "orderedNotes");
      assertEqual(
        result.titles,
        [
          "Chapter 5",
          "Frankenstein Lecture Series",
          "Chapter 4",
          "Chapter 2",
          "Chapter 3",
        ],
        "slide order should follow presentation relationships rather than filenames",
      );
      assertEqual(
        result.notes,
        [
          "Key themes:\ncreation, responsibility, isolation\n\nVisible date: 2026-03-12",
          "Compare to Prometheus myth",
        ],
        "notes should preserve line breaks, visible fields and intentional blank paragraphs in logical slide order",
      );
      assertEqual(
        result.orderedDiff,
        { added: [], changed: [], removed: [] },
        "reading existing slide notes should not mark package mutations",
      );
      assertBytesEqual(
        result.orderedBytesAfterRead,
        result.orderedBytes,
        "reading existing slide notes should keep archive bytes untouched",
      );
      assertEqual(result.minimalError.code, "PPTX_NOTES_MISSING", "missing notes should refuse with a stable code");
      assertBytesEqual(
        result.minimalBytesAfterRefusal,
        result.minimalSource,
        "reading notes on a slide without notes must not mutate the archive",
      );
      if (result.minimalNames.some((name) => name.startsWith("ppt/notesSlides/"))) {
        throw new Error("reading notes created a notes part unexpectedly");
      }
    },
  },
  {
    pattern: /^PPTX line-break and field text fixture is prepared from a real template$/,
    run: (context) => {
      const state = scenarioState(context);
      state.readableUnsupportedFixture = prepareReadableUnsupportedFixture();
      state.readableUnsupported = undefined;
    },
  },
  {
    pattern: /^PPTX inspects the paragraph text and attempts an anchored edit on that topology$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.readableUnsupported = await runReadableUnsupportedScenario(
        required(state.readableUnsupportedFixture, "readableUnsupportedFixture"),
      );
    },
  },
  {
    pattern: /^PPTX exposes line breaks and field text faithfully and refuses the unsupported edit without mutation$/,
    run: (context) => {
      const state = scenarioState(context);
      const result = required(state.readableUnsupported, "readableUnsupported");
      assertEqual(result.paragraphText, "Chapter\n7 Notes", "inspectText should preserve line breaks and visible field text");
      assertEqual(
        result.paragraphRuns,
        [
          { text: "Chapter", attrs: { b: "1" } },
          { text: "\n", attrs: { lang: "en-US" } },
          { text: "7", attrs: { i: "1" } },
          { text: " Notes", attrs: { u: "sng" } },
        ],
        "inspectText runs should expose readable fragments in order",
      );
      assertEqual(
        result.error.code,
        "PPTX_UNSUPPORTED_TEXT_TOPOLOGY",
        "editing a paragraph with breaks or fields should refuse with a stable code",
      );
      assertEqual(
        result.diffAfterRefusal,
        { added: [], changed: [], removed: [] },
        "unsupported edit refusal should leave the package diff empty",
      );
      assertBytesEqual(
        result.bytesAfterRefusal,
        result.sourceBytes,
        "unsupported edit refusal should leave package bytes untouched",
      );
    },
  },
  {
    pattern: /^PPTX fragmented title fixture is prepared from a real template and an untouched ZIP member$/,
    run: (context) => {
      const state = scenarioState(context);
      state.crossRunFixture = prepareCrossRunFixture();
      state.crossRun = undefined;
    },
  },
  {
    pattern: /^PPTX replaces anchored cross-run text and saves then reopens the package$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.crossRun = await runCrossRunReplacementScenario(required(state.crossRunFixture, "crossRunFixture"));
    },
  },
  {
    pattern: /^PPTX preserves the replacement text, the starting run formatting, and unrelated ZIP member bytes$/,
    run: (context) => {
      const state = scenarioState(context);
      const result = required(state.crossRun, "crossRun");
      assertEqual(result.reopenedTitle, "Friendstein", "replacement should round-trip through save/reopen");
      assertEqual(
        result.changedParts,
        { added: [], changed: ["ppt/slides/slide1.xml"], removed: [] },
        "only the edited slide part should change",
      );
      assertBytesEqual(result.customMemberBytes, encoder.encode("keep-me-safe"), "custom member bytes should stay untouched");
      assertBytesEqual(
        result.docPropsBytes,
        readZip(result.sourceBytes).get("docProps/app.xml") ?? new Uint8Array(),
        "unrelated docProps bytes should stay untouched",
      );
      assertEqual(
        result.titleRuns,
        [
          { text: "Fr", attrs: { b: "1" } },
          { text: "iend", attrs: { b: "1" } },
          { text: "stein", attrs: { u: "sng" } },
        ],
        "cross-run replacement should keep boundary fragments and replacement formatting exact",
      );
    },
  },
  {
    pattern: /^PPTX stale-anchor fixture is prepared from a real template$/,
    run: (context) => {
      const state = scenarioState(context);
      state.staleAnchorFixture = prepareStaleAnchorFixture();
      state.staleAnchor = undefined;
    },
  },
  {
    pattern: /^PPTX replaces anchored text once and retries with the stale anchor$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.staleAnchor = await runStaleAnchorScenario(required(state.staleAnchorFixture, "staleAnchorFixture"));
    },
  },
  {
    pattern: /^PPTX refuses the stale anchor and keeps the post-success bytes unchanged$/,
    run: (context) => {
      const state = scenarioState(context);
      const result = required(state.staleAnchor, "staleAnchor");
      assertEqual(result.error.code, "PPTX_STALE_ANCHOR", "stale anchors should refuse with a stable code");
      assertBytesEqual(
        result.bytesAfterRefusal,
        result.bytesAfterSuccess,
        "stale anchor refusal should not mutate the package after a successful edit",
      );
      assertEqual(result.reopenedTitle, "Creature", "post-refusal reopen should keep the successful replacement only");
    },
  },
];

function scenarioState(context: Record<string, unknown>): AcceptanceState {
  return context.state as AcceptanceState;
}

function buildRelationshipOrderedNotesFixture(): Uint8Array {
  const parts = readZip(readFixtureBytesSync(GO_NOTES_FIXTURE));
  const presentationXml = decodeRequiredPart(parts, "ppt/presentation.xml");
  const matches = presentationXml.match(/<p:sldId\b[^>]*\/>/g);
  if (!matches || matches.length < 5) {
    throw new Error("expected five slide ids in notes fixture");
  }

  const reordered = [matches[4], matches[0], matches[3], matches[1], matches[2]].join("");
  const rewrittenPresentation = presentationXml.replace(matches.join(""), reordered);
  if (rewrittenPresentation === presentationXml) {
    throw new Error("failed to reorder presentation slide ids");
  }

  const notesXml = decodeRequiredPart(parts, "ppt/notesSlides/notesSlide5.xml");
  const originalNotesParagraph = '<a:p><a:r><a:t>Key themes: creation, responsibility, isolation</a:t></a:r></a:p>';
  if (!notesXml.includes(originalNotesParagraph)) {
    throw new Error("unexpected notes slide structure in go fixture");
  }

  parts.set("ppt/presentation.xml", encoder.encode(rewrittenPresentation));
  parts.set(
    "ppt/notesSlides/notesSlide5.xml",
    encoder.encode(notesXml.replace(originalNotesParagraph, COMPLEX_NOTES_PARAGRAPHS)),
  );
  return writeZip(parts);
}

function buildReadableUnsupportedSlideFixture(): Uint8Array {
  const parts = readZip(readFixtureBytesSync(PYTHON_TITLE_FIXTURE));
  const slideXml = decodeRequiredPart(parts, "ppt/slides/slide1.xml");
  const originalParagraph = '<a:p><a:r><a:t>Frankenstein</a:t></a:r></a:p>';
  if (!slideXml.includes(originalParagraph)) {
    throw new Error("unexpected title slide structure in python fixture");
  }

  parts.set(
    "ppt/slides/slide1.xml",
    encoder.encode(slideXml.replace(originalParagraph, BREAK_FIELD_TITLE_PARAGRAPH)),
  );
  return writeZip(parts);
}

function buildFragmentedTitleFixture(): Uint8Array {
  const parts = readZip(readFixtureBytesSync(PYTHON_TITLE_FIXTURE));
  const slideXml = decodeRequiredPart(parts, "ppt/slides/slide1.xml");
  const originalParagraph = '<a:p><a:r><a:t>Frankenstein</a:t></a:r></a:p>';
  if (!slideXml.includes(originalParagraph)) {
    throw new Error("unexpected title slide structure in python fixture");
  }

  parts.set(
    "ppt/slides/slide1.xml",
    encoder.encode(slideXml.replace(originalParagraph, FRAGMENTED_TITLE_PARAGRAPH)),
  );
  parts.set("custom/data.bin", encoder.encode("keep-me-safe"));

  const contentTypesXml = decodeRequiredPart(parts, "[Content_Types].xml");
  if (!contentTypesXml.includes('Extension="bin"')) {
    parts.set(
      "[Content_Types].xml",
      encoder.encode(
        contentTypesXml.replace(
          "</Types>",
          '<Default Extension="bin" ContentType="application/octet-stream"/></Types>',
        ),
      ),
    );
  }

  return writeZip(parts);
}

function readFirstParagraphRuns(xml: string): Array<{ text: string; attrs: Record<string, string> }> {
  const doc = parseXml(xml);
  const paragraph = firstParagraph(doc.root);
  if (!paragraph) {
    throw new Error("expected a paragraph in the slide XML");
  }

  return paragraph.children
    .filter((child) => child.localName === "r" && child.namespaceURI === DRAWING_NS)
    .map((run) => {
      const textElement = run.children.find(
        (child) => child.localName === "t" && child.namespaceURI === DRAWING_NS,
      );
      const rPr = run.children.find(
        (child) => child.localName === "rPr" && child.namespaceURI === DRAWING_NS,
      );
      return {
        text: textElement?.text ?? "",
        attrs: { ...(rPr?.attributes ?? {}) },
      };
    });
}

function firstParagraph(root: ReturnType<typeof parseXml>["root"]): ReturnType<typeof parseXml>["root"] | undefined {
  const stack = [root];
  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) {
      continue;
    }
    if (current.localName === "p" && current.namespaceURI === DRAWING_NS) {
      return current;
    }
    for (let index = current.children.length - 1; index >= 0; index -= 1) {
      const child = current.children[index];
      if (child) {
        stack.push(child);
      }
    }
  }
  return undefined;
}

async function saveToTempAndReadBytes(presentation: Presentation): Promise<Uint8Array> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-pptx-"));
  try {
    const path = join(root, "roundtrip.pptx");
    await presentation.save(path);
    return Uint8Array.from(await Bun.file(path).bytes());
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}

function decodeRequiredPart(parts: ReadonlyMap<string, Uint8Array>, name: string): string {
  const part = parts.get(name);
  if (!part) {
    throw new Error(`missing ZIP member ${name}`);
  }
  return decoder.decode(part);
}

function readFixtureBytesSync(path: string): Uint8Array {
  return Uint8Array.from(readFileSync(path));
}

function assertEqual<T>(actual: T, expected: T, message: string): void {
  const left = JSON.stringify(actual);
  const right = JSON.stringify(expected);
  if (left !== right) {
    throw new Error(`${message}\nexpected: ${right}\nactual:   ${left}`);
  }
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

function required<T>(value: T | undefined, label: string): T {
  assert.ok(value !== undefined, `missing ${label}`);
  return value;
}

const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
