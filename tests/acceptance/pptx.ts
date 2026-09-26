import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { Presentation, type TextAnchor } from "../../src/pptx/index.ts";
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

type OrderedNotesResult = {
  titles: string[];
  notes: string[];
  orderedBytes: Uint8Array;
  minimalError: OoxmlError;
  minimalBytesAfterRefusal: Uint8Array;
  minimalNames: string[];
};

type CrossRunResult = {
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
  orderedNotes?: OrderedNotesResult;
  crossRun?: CrossRunResult;
  staleAnchor?: StaleAnchorResult;
};

export async function runOrderedNotesScenario(): Promise<OrderedNotesResult> {
  const orderedBytes = buildRelationshipOrderedNotesFixture();
  const orderedPresentation = await Presentation.open(orderedBytes);
  const titles = orderedPresentation.slides.map((slide) => {
    const paragraphs = slide.inspectText("acceptance.order.notes");
    return paragraphs[0]?.text ?? "";
  });
  const notes = [
    orderedPresentation.slides[0]?.readNotesText() ?? "",
    orderedPresentation.slides.at(-1)?.readNotesText() ?? "",
  ];

  const minimalSource = Uint8Array.from(await Bun.file(GO_MINIMAL_FIXTURE).bytes());
  const minimalPresentation = await Presentation.open(minimalSource);
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

  const minimalBytesAfterRefusal = minimalPresentation.package.toBytes();
  const minimalNames = minimalPresentation.package.names();

  assertEqual(
    titles,
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
    notes,
    [
      "Key themes: creation, responsibility, isolation",
      "Compare to Prometheus myth",
    ],
    "notes should follow reordered logical slide order",
  );
  assertEqual(orderedPresentation.package.diff(), {
    added: [],
    changed: [],
    removed: [],
  }, "reading existing slide notes should not mark package mutations");
  assertBytesEqual(
    orderedPresentation.package.toBytes(),
    orderedBytes,
    "reading existing slide notes should keep archive bytes untouched",
  );
  assertEqual(minimalError.code, "PPTX_NOTES_MISSING", "missing notes should refuse with a stable code");
  assertBytesEqual(
    minimalBytesAfterRefusal,
    minimalSource,
    "reading notes on a slide without notes must not mutate the archive",
  );
  if (minimalNames.some((name) => name.startsWith("ppt/notesSlides/"))) {
    throw new Error("reading notes created a notes part unexpectedly");
  }

  return {
    titles,
    notes,
    orderedBytes,
    minimalError,
    minimalBytesAfterRefusal,
    minimalNames,
  };
}

export async function runCrossRunReplacementScenario(): Promise<CrossRunResult> {
  const source = buildFragmentedTitleFixture();
  const presentation = await Presentation.open(source);
  const slide = presentation.slides[0];
  if (!slide) {
    throw new Error("fragmented fixture should contain slide 1");
  }

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

  const titleRuns = readFirstParagraphRuns(decoder.decode(savedParts.get("ppt/slides/slide1.xml") ?? new Uint8Array()));

  assertEqual(reopenedTitle, "Friendstein", "replacement should round-trip through save/reopen");
  assertEqual(changedParts, {
    added: [],
    changed: ["ppt/slides/slide1.xml"],
    removed: [],
  }, "only the edited slide part should change");
  assertBytesEqual(customMemberBytes, encoder.encode("keep-me-safe"), "custom member bytes should stay untouched");
  assertBytesEqual(
    docPropsBytes,
    readZip(source).get("docProps/app.xml") ?? new Uint8Array(),
    "unrelated docProps bytes should stay untouched",
  );
  assertEqual(
    titleRuns.map((run) => ({ text: run.text, attrs: run.attrs })),
    [
      { text: "Fr", attrs: { b: "1" } },
      { text: "iend", attrs: { b: "1" } },
      { text: "stein", attrs: { u: "sng" } },
    ],
    "cross-run replacement should keep boundary fragments and replacement formatting exact",
  );

  return {
    reopenedTitle,
    changedParts,
    customMemberBytes,
    docPropsBytes,
    titleRuns,
  };
}

export async function runStaleAnchorScenario(): Promise<StaleAnchorResult> {
  const source = Uint8Array.from(await Bun.file(PYTHON_TITLE_FIXTURE).bytes());
  const presentation = await Presentation.open(source);
  const slide = presentation.slides[0];
  if (!slide) {
    throw new Error("stale-anchor fixture should contain slide 1");
  }

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
  const reopenedTitle = reopened.slides[0]?.inspectText("acceptance.stale.anchor.reopen")[0]?.text ?? "";

  assertEqual(refusal.code, "PPTX_STALE_ANCHOR", "stale anchors should refuse with a stable code");
  assertBytesEqual(
    bytesAfterRefusal,
    bytesAfterSuccess,
    "stale anchor refusal should not mutate the package after a successful edit",
  );
  assertEqual(reopenedTitle, "Creature", "post-refusal reopen should keep the successful replacement only");

  return {
    error: refusal,
    bytesAfterSuccess,
    bytesAfterRefusal,
    reopenedTitle,
  };
}

export const bindings: StepBinding[] = [
  {
    pattern: /^PPTX ordered notes fixtures are prepared$/,
    run: (context) => {
      const state = context.state as AcceptanceState;
      state.orderedNotes = undefined;
    },
  },
  {
    pattern: /^PPTX opens the reordered notes fixture and probes notes reads$/,
    run: async (context) => {
      const state = context.state as AcceptanceState;
      state.orderedNotes = await runOrderedNotesScenario();
    },
  },
  {
    pattern: /^PPTX keeps slide order and notes reads non-mutating without creating missing notes parts$/,
    run: (context) => {
      const state = context.state as AcceptanceState;
      if (!state.orderedNotes) {
        throw new Error("ordered notes scenario did not run");
      }
    },
  },
  {
    pattern: /^PPTX fragmented title fixture is prepared from a real template and an untouched ZIP member$/,
    run: (context) => {
      const state = context.state as AcceptanceState;
      state.crossRun = undefined;
    },
  },
  {
    pattern: /^PPTX replaces anchored cross-run text and saves then reopens the package$/,
    run: async (context) => {
      const state = context.state as AcceptanceState;
      state.crossRun = await runCrossRunReplacementScenario();
    },
  },
  {
    pattern: /^PPTX preserves the replacement text, the starting run formatting, and unrelated ZIP member bytes$/,
    run: (context) => {
      const state = context.state as AcceptanceState;
      if (!state.crossRun) {
        throw new Error("cross-run scenario did not run");
      }
    },
  },
  {
    pattern: /^PPTX stale-anchor fixture is prepared from a real template$/,
    run: (context) => {
      const state = context.state as AcceptanceState;
      state.staleAnchor = undefined;
    },
  },
  {
    pattern: /^PPTX replaces anchored text once and retries with the stale anchor$/,
    run: async (context) => {
      const state = context.state as AcceptanceState;
      state.staleAnchor = await runStaleAnchorScenario();
    },
  },
  {
    pattern: /^PPTX refuses the stale anchor and keeps the post-success bytes unchanged$/,
    run: (context) => {
      const state = context.state as AcceptanceState;
      if (!state.staleAnchor) {
        throw new Error("stale anchor scenario did not run");
      }
    },
  },
];

function buildRelationshipOrderedNotesFixture(): Uint8Array {
  const parts = readZip(readFixtureBytesSync(GO_NOTES_FIXTURE));
  const presentationXml = decodeRequiredPart(parts, "ppt/presentation.xml");
  const matches = presentationXml.match(/<p:sldId\b[^>]*\/>/g);
  if (!matches || matches.length < 5) {
    throw new Error("expected five slide ids in notes fixture");
  }

  const reordered = [matches[4], matches[0], matches[3], matches[1], matches[2]].join("");
  const rewritten = presentationXml.replace(matches.join(""), reordered);
  if (rewritten === presentationXml) {
    throw new Error("failed to reorder presentation slide ids");
  }

  parts.set("ppt/presentation.xml", encoder.encode(rewritten));
  return writeZip(parts);
}

function buildFragmentedTitleFixture(): Uint8Array {
  const parts = readZip(readFixtureBytesSync(PYTHON_TITLE_FIXTURE));
  const slideXml = decodeRequiredPart(parts, "ppt/slides/slide1.xml");
  const originalParagraph = "<a:p><a:r><a:t>Frankenstein</a:t></a:r></a:p>";
  if (!slideXml.includes(originalParagraph)) {
    throw new Error("unexpected title slide structure in python fixture");
  }

  parts.set(
    "ppt/slides/slide1.xml",
    encoder.encode(slideXml.replace(originalParagraph, FRAGMENTED_TITLE_PARAGRAPH)),
  );
  parts.set("custom/data.bin", encoder.encode("keep-me-safe"));

  const contentTypesXml = decodeRequiredPart(parts, "[Content_Types].xml");
  if (!contentTypesXml.includes("Extension=\"bin\"")) {
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

const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
