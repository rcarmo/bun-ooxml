import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { Presentation, type TableGeometry } from "../../src/pptx/index.ts";
import { parseXml, type XmlElement } from "../../src/xml/index.ts";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const TABLE_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";
const SLIDE_PART = "ppt/slides/slide1.xml";
const DEFAULT_GEOMETRY: TableGeometry = { x: 120, y: 240, width: 1001, height: 1003 };
const STALE_FIRST_GEOMETRY: TableGeometry = { x: 100, y: 100, width: 900, height: 600 };
const STALE_SECOND_GEOMETRY: TableGeometry = { x: 1200, y: 100, width: 900, height: 600 };
const DEFAULT_EMPTY_CELL_XML = '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p></a:p></a:txBody><a:tcPr/></a:tc>';
const STYLED_CELL_XML = [
  '<a:tc>',
  '<a:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:pPr algn="r"/><a:r><a:rPr b="1" sz="1800"/><a:t>Galvanic battery</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></a:txBody>',
  '<a:tcPr marL="111"><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:tcPr>',
  '</a:tc>',
].join("");

type GeometryRoundTripResult = {
  savedBytes: Uint8Array;
  reopenedRows: number;
  reopenedColumns: number;
  reopenedText: string;
  metrics: TableMetrics;
};

type FormattingPreservationResult = {
  savedBytes: Uint8Array;
  reopenedText: string;
  formatting: TableCellFormatting;
};

type StaleTableHandleResult = {
  error: OoxmlError;
  bytesAfterMutation: Uint8Array;
  bytesAfterRefusal: Uint8Array;
};

type AtomicRefusalResult = {
  sourceBytes: Uint8Array;
  error: OoxmlError;
  bytesAfterRefusal: Uint8Array;
  diffAfterRefusal: { added: string[]; changed: string[]; removed: string[] };
};

type TableMetrics = {
  x: number;
  y: number;
  width: number;
  height: number;
  columnWidths: number[];
  rowHeights: number[];
};

type TableCellFormatting = {
  bodyPrAttrs: Record<string, string>;
  tcPrAttrs: Record<string, string>;
  fillColor: string;
  paragraphAttrs: Record<string, string>;
  firstRunAttrs: Record<string, string>;
  endParaRPrAttrs: Record<string, string>;
};

type AcceptanceState = {
  geometryRoundTrip?: GeometryRoundTripResult;
  formattingPreservation?: FormattingPreservationResult;
  staleTableHandle?: StaleTableHandleResult;
  atomicRefusal?: AtomicRefusalResult;
};

export async function runGeometryRoundTripScenario(): Promise<GeometryRoundTripResult> {
  const presentation = Presentation.create();
  const slide = presentation.addTextSlide("Tables");
  slide.addTable(2, 3, DEFAULT_GEOMETRY).cell(0, 0).text = "  <Alpha & Beta>  ";

  const savedBytes = await saveToTempAndReadBytes(presentation);
  const reopened = await Presentation.open(savedBytes);
  const reopenedTable = required(reopened.slides[0]?.tables[0], "reopened slide 1 table 1");

  return {
    savedBytes,
    reopenedRows: reopenedTable.rows,
    reopenedColumns: reopenedTable.columns,
    reopenedText: reopenedTable.cell(0, 0).text,
    metrics: readFirstTableMetrics(savedBytes),
  };
}

export function prepareFormattingPreservationFixture(): Uint8Array {
  const parts = readZip(buildBaseTableFixture({ rows: 1, columns: 1, geometry: { x: 500, y: 700, width: 2000, height: 1200 } }));
  const slideXml = decodeRequiredPart(parts, SLIDE_PART);
  parts.set(SLIDE_PART, encoder.encode(replaceRequired(slideXml, DEFAULT_EMPTY_CELL_XML, STYLED_CELL_XML, "styled table cell")));
  return writeZip(parts);
}

export async function runFormattingPreservationScenario(
  source: Uint8Array = prepareFormattingPreservationFixture(),
): Promise<FormattingPreservationResult> {
  const presentation = await Presentation.open(source);
  const slide = required(presentation.slides[0], "styled fixture slide 1");
  required(slide.tables[0], "styled fixture table 1").cell(0, 0).text = "updated value";

  const savedBytes = await saveToTempAndReadBytes(presentation);
  const reopened = await Presentation.open(savedBytes);
  const reopenedCell = required(reopened.slides[0]?.tables[0], "reopened styled table").cell(0, 0);

  return {
    savedBytes,
    reopenedText: reopenedCell.text,
    formatting: readFirstTableCellFormatting(savedBytes),
  };
}

export async function runStaleTableHandleScenario(): Promise<StaleTableHandleResult> {
  const presentation = Presentation.create();
  const slide = presentation.addTextSlide("Tables");
  const staleCell = slide.addTable(1, 1, STALE_FIRST_GEOMETRY).cell(0, 0);
  slide.addTable(1, 1, STALE_SECOND_GEOMETRY);
  const bytesAfterMutation = presentation.package.toBytes();

  let error: OoxmlError | undefined;
  try {
    staleCell.text = "stale write";
  } catch (caught) {
    if (caught instanceof OoxmlError) {
      error = caught;
    } else {
      throw caught;
    }
  }
  if (!error) {
    throw new Error("expected stale PPTX table handle refusal");
  }

  return {
    error,
    bytesAfterMutation,
    bytesAfterRefusal: presentation.package.toBytes(),
  };
}

export function prepareMergedTableFixture(): Uint8Array {
  return rewriteFirstSlideXml(
    buildBaseTableFixture({ rows: 2, columns: 2, geometry: { x: 300, y: 400, width: 1800, height: 900 } }),
    (xml) => replaceRequired(xml, "<a:tc><a:txBody>", '<a:tc gridSpan="2"><a:txBody>', "merged table cell"),
  );
}

export function prepareMalformedMergeFixture(): Uint8Array {
  return rewriteFirstSlideXml(
    buildBaseTableFixture({ rows: 2, columns: 2, geometry: { x: 300, y: 400, width: 1800, height: 900 } }),
    (xml) => replaceRequired(xml, "<a:tc><a:txBody>", '<a:tc hMerge="maybe"><a:txBody>', "malformed merge cell"),
  );
}

export async function runAtomicRefusalScenario(kind: "merged-cell" | "malformed-merge"): Promise<AtomicRefusalResult> {
  const sourceBytes = kind === "merged-cell" ? prepareMergedTableFixture() : prepareMalformedMergeFixture();
  const presentation = await Presentation.open(sourceBytes);

  let error: OoxmlError | undefined;
  try {
    runRefusalCase(presentation, kind);
  } catch (caught) {
    if (caught instanceof OoxmlError) {
      error = caught;
    } else {
      throw caught;
    }
  }
  if (!error) {
    throw new Error(`expected PPTX table refusal for ${kind}`);
  }

  return {
    sourceBytes,
    error,
    bytesAfterRefusal: presentation.package.toBytes(),
    diffAfterRefusal: presentation.package.diff(),
  };
}

export const bindings: StepBinding[] = [
  {
    pattern: /^PPTX table round-trip scenario is prepared from a new presentation$/,
    run: (context) => {
      const state = scenarioState(context);
      state.geometryRoundTrip = undefined;
    },
  },
  {
    pattern: /^PPTX authors a rectangular table and saves then reopens it$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.geometryRoundTrip = await runGeometryRoundTripScenario();
    },
  },
  {
    pattern: /^PPTX preserves the table size, exact grid sums, and cell text after reopen$/,
    run: (context) => {
      const result = required(scenarioState(context).geometryRoundTrip, "geometryRoundTrip");
      assertEqual(result.reopenedRows, 2, "reopened table should preserve row count");
      assertEqual(result.reopenedColumns, 3, "reopened table should preserve column count");
      assertEqual(result.reopenedText, "  <Alpha & Beta>  ", "reopened table should preserve authored cell text");
      assertEqual(result.metrics.x, DEFAULT_GEOMETRY.x, "table x should round-trip");
      assertEqual(result.metrics.y, DEFAULT_GEOMETRY.y, "table y should round-trip");
      assertEqual(result.metrics.width, DEFAULT_GEOMETRY.width, "table width should round-trip");
      assertEqual(result.metrics.height, DEFAULT_GEOMETRY.height, "table height should round-trip");
      assertEqual(sum(result.metrics.columnWidths), DEFAULT_GEOMETRY.width, "grid column widths should sum to table width");
      assertEqual(sum(result.metrics.rowHeights), DEFAULT_GEOMETRY.height, "row heights should sum to table height");
    },
  },
  {
    pattern: /^PPTX styled table fixture is prepared$/,
    run: (context) => {
      const state = scenarioState(context);
      state.formattingPreservation = undefined;
    },
  },
  {
    pattern: /^PPTX replaces a styled table cell and saves then reopens the presentation$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.formattingPreservation = await runFormattingPreservationScenario();
    },
  },
  {
    pattern: /^PPTX preserves table cell formatting across the update and reopen$/,
    run: (context) => {
      const result = required(scenarioState(context).formattingPreservation, "formattingPreservation");
      assertEqual(result.reopenedText, "updated value", "styled cell text should round-trip");
      assertEqual(result.formatting.bodyPrAttrs.wrap, "square", "bodyPr attributes should be preserved");
      assertEqual(result.formatting.tcPrAttrs.marL, "111", "tcPr attributes should be preserved");
      assertEqual(result.formatting.fillColor, "FFFF00", "tcPr child formatting should be preserved");
      assertEqual(result.formatting.paragraphAttrs.algn, "r", "paragraph properties should be preserved");
      assertEqual(result.formatting.firstRunAttrs, { b: "1", sz: "1800" }, "first run formatting should be preserved");
      assertEqual(result.formatting.endParaRPrAttrs.lang, "en-US", "end paragraph formatting should be preserved");
    },
  },
  {
    pattern: /^PPTX stale table handle scenario is prepared from a new presentation$/,
    run: (context) => {
      const state = scenarioState(context);
      state.staleTableHandle = undefined;
    },
  },
  {
    pattern: /^PPTX mutates the slide and retries a stale table cell handle$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.staleTableHandle = await runStaleTableHandleScenario();
    },
  },
  {
    pattern: /^PPTX refuses the stale table handle without mutating the package$/,
    run: (context) => {
      const result = required(scenarioState(context).staleTableHandle, "staleTableHandle");
      assertEqual(result.error.code, "PPTX_STALE_TABLE_HANDLE", "stale table handles should refuse with a stable code");
      assertBytesEqual(
        result.bytesAfterRefusal,
        result.bytesAfterMutation,
        "stale table refusal should leave package bytes unchanged",
      );
    },
  },
  {
    pattern: /^PPTX table refusal scenario "([^"]+)" is prepared$/,
    run: (context) => {
      const state = scenarioState(context);
      state.atomicRefusal = undefined;
    },
  },
  {
    pattern: /^PPTX attempts the table refusal "([^"]+)"$/,
    run: async (context, kind) => {
      const state = scenarioState(context);
      if (kind !== "merged-cell" && kind !== "malformed-merge") {
        throw new Error(`Unknown PPTX table refusal ${kind}`);
      }
      state.atomicRefusal = await runAtomicRefusalScenario(kind);
    },
  },
  {
    pattern: /^PPTX refusal "([^"]+)" is returned atomically for table refusal "([^"]+)"$/,
    run: (context, code, _kind) => {
      const result = required(scenarioState(context).atomicRefusal, "atomicRefusal");
      assertEqual(result.error.code, code, "table refusal should expose the expected stable code");
      assertEqual(result.diffAfterRefusal, { added: [], changed: [], removed: [] }, "refusal should keep the package diff empty");
      assertBytesEqual(result.bytesAfterRefusal, result.sourceBytes, "refusal should leave package bytes unchanged");
    },
  },
];

function scenarioState(context: Record<string, unknown>): AcceptanceState {
  return context.state as AcceptanceState;
}

function buildBaseTableFixture(options: { rows: number; columns: number; geometry: TableGeometry }): Uint8Array {
  const presentation = Presentation.create();
  const slide = presentation.addTextSlide("Tables");
  slide.addTable(options.rows, options.columns, options.geometry);
  return presentation.package.toBytes();
}

function rewriteFirstSlideXml(bytes: Uint8Array, rewrite: (xml: string) => string): Uint8Array {
  const parts = readZip(bytes);
  const original = decodeRequiredPart(parts, SLIDE_PART);
  const updated = rewrite(original);
  if (updated === original) {
    throw new Error("slide rewrite did not change the first slide XML");
  }
  parts.set(SLIDE_PART, encoder.encode(updated));
  return writeZip(parts);
}

function runRefusalCase(presentation: Presentation, kind: "merged-cell" | "malformed-merge"): void {
  const slide = required(presentation.slides[0], "slide 1");
  switch (kind) {
    case "merged-cell":
      required(slide.tables[0], "slide 1 table 1").cell(0, 0).text = "x";
      return;
    case "malformed-merge":
      void slide.tables[0];
      return;
  }
}

function readFirstTableMetrics(source: Uint8Array): TableMetrics {
  const slideXml = decodeRequiredPart(readZip(source), SLIDE_PART);
  const { frame, table } = firstSlideTable(slideXml);
  const xfrm = requiredChild(frame, "xfrm", PRESENTATION_NS);
  const off = requiredChild(xfrm, "off", DRAWING_NS);
  const ext = requiredChild(xfrm, "ext", DRAWING_NS);
  const grid = requiredChild(table, "tblGrid", DRAWING_NS);
  const rows = table.children.filter((child) => isElement(child, "tr", DRAWING_NS));

  return {
    x: parseRequiredInt(off.attributes.x, "table x"),
    y: parseRequiredInt(off.attributes.y, "table y"),
    width: parseRequiredInt(ext.attributes.cx, "table width"),
    height: parseRequiredInt(ext.attributes.cy, "table height"),
    columnWidths: grid.children
      .filter((child) => isElement(child, "gridCol", DRAWING_NS))
      .map((child) => parseRequiredInt(child.attributes.w, "grid column width")),
    rowHeights: rows.map((row) => parseRequiredInt(row.attributes.h, "row height")),
  };
}

function readFirstTableCellFormatting(source: Uint8Array): TableCellFormatting {
  const slideXml = decodeRequiredPart(readZip(source), SLIDE_PART);
  const { table } = firstSlideTable(slideXml);
  const row = required(
    table.children.find((child) => isElement(child, "tr", DRAWING_NS)),
    "first table row",
  );
  const cell = required(
    row.children.find((child) => isElement(child, "tc", DRAWING_NS)),
    "first table cell",
  );
  const txBody = requiredChild(cell, "txBody", DRAWING_NS);
  const bodyPr = requiredChild(txBody, "bodyPr", DRAWING_NS);
  const paragraph = requiredChild(txBody, "p", DRAWING_NS);
  const pPr = requiredChild(paragraph, "pPr", DRAWING_NS);
  const run = requiredChild(paragraph, "r", DRAWING_NS);
  const rPr = requiredChild(run, "rPr", DRAWING_NS);
  const endParaRPr = requiredChild(paragraph, "endParaRPr", DRAWING_NS);
  const tcPr = requiredChild(cell, "tcPr", DRAWING_NS);
  const solidFill = requiredChild(tcPr, "solidFill", DRAWING_NS);
  const srgbClr = requiredChild(solidFill, "srgbClr", DRAWING_NS);

  return {
    bodyPrAttrs: { ...bodyPr.attributes },
    tcPrAttrs: { ...tcPr.attributes },
    fillColor: required(srgbClr.attributes.val, "solid fill color"),
    paragraphAttrs: { ...pPr.attributes },
    firstRunAttrs: { ...rPr.attributes },
    endParaRPrAttrs: { ...endParaRPr.attributes },
  };
}

function firstSlideTable(slideXml: string): { frame: XmlElement; table: XmlElement } {
  const parsed = parseXml(slideXml);
  const frame = required(findTableGraphicFrame(parsed.root), "slide table graphic frame");
  const graphic = requiredChild(frame, "graphic", DRAWING_NS);
  const graphicData = requiredChild(graphic, "graphicData", DRAWING_NS);
  if (graphicData.attributes.uri !== TABLE_GRAPHIC_DATA_URI) {
    throw new Error(`expected table graphicData URI, got ${JSON.stringify(graphicData.attributes.uri)}`);
  }
  return {
    frame,
    table: requiredChild(graphicData, "tbl", DRAWING_NS),
  };
}

function findTableGraphicFrame(root: XmlElement): XmlElement | undefined {
  const stack = [root];
  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) {
      continue;
    }
    if (isElement(current, "graphicFrame", PRESENTATION_NS)) {
      const graphic = current.children.find((child) => isElement(child, "graphic", DRAWING_NS));
      const graphicData = graphic?.children.find((child) => isElement(child, "graphicData", DRAWING_NS));
      const table = graphicData?.children.find((child) => isElement(child, "tbl", DRAWING_NS));
      if (graphicData?.attributes.uri === TABLE_GRAPHIC_DATA_URI && table) {
        return current;
      }
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

function requiredChild(parent: XmlElement, localName: string, namespaceURI: string): XmlElement {
  return required(
    parent.children.find((child) => isElement(child, localName, namespaceURI)),
    `${parent.name}/${localName}`,
  );
}

function parseRequiredInt(raw: string | undefined, label: string): number {
  const value = raw === undefined ? Number.NaN : Number.parseInt(raw, 10);
  if (!Number.isSafeInteger(value)) {
    throw new Error(`invalid ${label}: ${JSON.stringify(raw)}`);
  }
  return value;
}

function isElement(element: XmlElement, localName: string, namespaceURI: string): boolean {
  return element.localName === localName && element.namespaceURI === namespaceURI;
}

async function saveToTempAndReadBytes(presentation: Presentation): Promise<Uint8Array> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-pptx-tables-"));
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

function replaceRequired(text: string, from: string, to: string, label: string): string {
  if (!text.includes(from)) {
    throw new Error(`missing ${label} marker`);
  }
  const updated = text.replace(from, to);
  if (updated === text) {
    throw new Error(`failed to rewrite ${label}`);
  }
  return updated;
}

function sum(values: readonly number[]): number {
  return values.reduce((total, value) => total + value, 0);
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
