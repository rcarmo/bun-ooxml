import { expect } from "bun:test";
import { readFileSync } from "node:fs";
import { join, resolve } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { Document, W_NS, type Table, type TableCell } from "../../src/docx/index.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { applyEdits, parseXml, type XmlElement } from "../../src/xml/index.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const PYTHON_WORD_ROOT = join(
  PROJECT_ROOT,
  "fixtures/python-office-mcp-server/tests/_templates/testdata/word",
);
const DOCUMENT_PART = "word/document.xml";
const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });
const UTF8_ENCODER = new TextEncoder();

type AcceptanceState = {
  document?: Document;
  rememberedBytes?: Uint8Array;
  rememberedCell?: TableCell;
  rememberedOpaquePart?: { name: string; bytes: Uint8Array };
  lastError?: unknown;
};

export const bindings: StepBinding[] = [
  {
    pattern: /^DOCX table source "([^"]+)" is prepared$/,
    run: async (context, source) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedBytes = undefined;
      state.rememberedCell = undefined;
      state.rememberedOpaquePart = undefined;
      state.document = source === "new-document"
        ? Document.create()
        : await Document.open(resolveFixture(source));
    },
  },
  {
    pattern: /^DOCX table (\d+)x(\d+) is appended$/,
    run: (context, rowsText, columnsText) => {
      const state = getState(context);
      state.lastError = undefined;
      document(state).addTable(Number.parseInt(rowsText, 10), Number.parseInt(columnsText, 10));
    },
  },
  {
    pattern: /^DOCX table (\d+) cell \((\d+),(\d+)\) text is set to "([^"]*)"$/,
    run: (context, tableText, rowText, columnText, rawText) => {
      const state = getState(context);
      state.lastError = undefined;
      table(state, tableText).cell(Number.parseInt(rowText, 10), Number.parseInt(columnText, 10)).text = decodeStepText(rawText);
    },
  },
  {
    pattern: /^DOCX table document is saved and reopened$/,
    run: async (context) => {
      const state = getState(context);
      state.lastError = undefined;
      state.document = await Document.open(await document(state).save());
    },
  },
  {
    pattern: /^DOCX table (\d+) has (\d+) rows and (\d+) columns$/,
    run: (context, tableText, rowsText, columnsText) => {
      const current = table(getState(context), tableText);
      expect(current.rows).toBe(Number.parseInt(rowsText, 10));
      expect(current.columns).toBe(Number.parseInt(columnsText, 10));
    },
  },
  {
    pattern: /^DOCX table (\d+) cell \((\d+),(\d+)\) text equals "([^"]*)"$/,
    run: (context, tableText, rowText, columnText, rawExpected) => {
      const current = table(getState(context), tableText);
      expect(current.cell(Number.parseInt(rowText, 10), Number.parseInt(columnText, 10)).text).toBe(
        decodeStepText(rawExpected),
      );
    },
  },
  {
    pattern: /^DOCX table (\d+) is stored before the section properties$/,
    run: (context, tableText) => {
      const body = readBody(document(getState(context)));
      const children = body.children.filter((child) => isWord(child, "tbl") || isWord(child, "sectPr"));
      const index = Number.parseInt(tableText, 10) - 1;
      expect(children[index]?.localName).toBe("tbl");
      expect(children[index + 1]?.localName).toBe("sectPr");
    },
  },
  {
    pattern: /^DOCX table (\d+) cell \((\d+),(\d+)\) XML preserves boundary spaces and escapes special characters$/,
    run: (context, tableText, rowText, columnText) => {
      const xml = rawCellXml(
        document(getState(context)),
        Number.parseInt(tableText, 10) - 1,
        Number.parseInt(rowText, 10),
        Number.parseInt(columnText, 10),
      );
      expect(xml).toContain('<w:t xml:space="preserve">  &lt;Alpha &amp; Beta&gt;  </w:t>');
    },
  },
  {
    pattern: /^DOCX table (\d+) cell \((\d+),(\d+)\) is remembered$/,
    run: (context, tableText, rowText, columnText) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedCell = table(state, tableText).cell(Number.parseInt(rowText, 10), Number.parseInt(columnText, 10));
    },
  },
  {
    pattern: /^DOCX table current saved bytes are remembered$/,
    run: async (context) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedBytes = await document(state).save();
    },
  },
  {
    pattern: /^DOCX table stale text set to "([^"]*)" is attempted on the remembered cell$/,
    run: (context, rawText) => {
      const state = getState(context);
      state.lastError = undefined;
      try {
        requireDefined(state.rememberedCell, "Missing remembered DOCX table cell").text = decodeStepText(rawText);
      } catch (error) {
        state.lastError = error;
        return;
      }
      throw new Error("Expected stale DOCX table cell write to fail");
    },
  },
  {
    pattern: /^DOCX table refusal "([^"]+)" is attempted$/,
    run: (context, refusalCase) => {
      const state = getState(context);
      state.lastError = undefined;
      try {
        runRefusalCase(document(state), refusalCase);
      } catch (error) {
        state.lastError = error;
        return;
      }
      throw new Error(`Expected DOCX table refusal for ${refusalCase}`);
    },
  },
  {
    pattern: /^DOCX table refusal code equals "([^"]+)"$/,
    run: (context, expectedCode) => {
      const state = getState(context);
      if (expectedCode === "range") {
        expect(state.lastError).toBeInstanceOf(RangeError);
        return;
      }
      expect(state.lastError).toBeInstanceOf(OoxmlError);
      expect((state.lastError as OoxmlError).code).toBe(expectedCode);
    },
  },
  {
    pattern: /^DOCX table saved bytes equal the remembered bytes$/,
    run: async (context) => {
      const state = getState(context);
      expect(await document(state).save()).toEqual(requireDefined(state.rememberedBytes, "Missing remembered DOCX bytes"));
    },
  },
  {
    pattern: /^DOCX table opaque part "([^"]+)" bytes are remembered$/,
    run: (context, partName) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedOpaquePart = {
        name: partName,
        bytes: requireDefined(document(state).package.get(partName), `Missing DOCX part ${partName}`),
      };
    },
  },
  {
    pattern: /^DOCX table opaque part "([^"]+)" bytes are unchanged$/,
    run: (context, partName) => {
      const state = getState(context);
      const remembered = requireDefined(state.rememberedOpaquePart, "Missing remembered opaque DOCX part");
      expect(remembered.name).toBe(partName);
      expect(document(state).package.get(partName)).toEqual(remembered.bytes);
    },
  },
];

function getState(context: Record<string, unknown>): AcceptanceState {
  return requireDefined(context.state as AcceptanceState | undefined, "Missing acceptance state");
}

function document(state: AcceptanceState): Document {
  return requireDefined(state.document, "Missing DOCX document");
}

function table(state: AcceptanceState, tableText: string): Table {
  const index = Number.parseInt(tableText, 10) - 1;
  return requireDefined(document(state).tables[index], `Missing DOCX table ${index + 1}`);
}

function runRefusalCase(doc: Document, refusalCase: string): void {
  switch (refusalCase) {
    case "row-oob":
      requireDefined(doc.tables[0], "Missing DOCX table 1").cell(9, 0).text = "x";
      return;
    case "merged-cell":
      requireDefined(doc.tables[0], "Missing DOCX table 1").cell(0, 0).text = "x";
      return;
    case "nested-cell":
      requireDefined(doc.tables[0], "Missing DOCX table 1").cell(2, 0).text = "x";
      return;
    case "bizarre":
      requireDefined(doc.tables[0], "Missing DOCX table 1").cell(0, 0).text = "x";
      return;
    default:
      throw new Error(`Unknown DOCX table refusal case ${refusalCase}`);
  }
}

function resolveFixture(fixture: string): string | Uint8Array {
  if (fixture === "synthetic-grid-before") {
    return buildSyntheticGridBeforeFixture();
  }
  return join(PROJECT_ROOT, fixture);
}

function buildSyntheticGridBeforeFixture(): Uint8Array {
  const source = new Uint8Array(readFileSync(join(PYTHON_WORD_ROOT, "simple_table.docx")));
  const parts = readZip(source);
  const xml = UTF8_DECODER.decode(requireDefined(parts.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`));
  const parsed = parseXml(xml);
  const body = requireDefined(parsed.root.children.find((child) => isWord(child, "body")), "Missing w:body");
  const tableElement = requireDefined(body.children.find((child) => isWord(child, "tbl")), "Missing first w:tbl");
  const firstRow = requireDefined(tableElement.children.find((child) => isWord(child, "tr")), "Missing first w:tr");
  const firstCell = requireDefined(firstRow.children.find((child) => isWord(child, "tc")), "Missing first w:tc");
  const updated = applyEdits(xml, [
    { start: firstCell.start, end: firstCell.end, value: "" },
    { start: firstRow.openEnd, end: firstRow.openEnd, value: '<w:trPr><w:gridBefore w:val="1"/></w:trPr>' },
  ]);
  parts.set(DOCUMENT_PART, UTF8_ENCODER.encode(updated));
  return writeZip(parts);
}

function rawCellXml(doc: Document, tableIndex: number, row: number, column: number): string {
  const tableElement = readTopLevelTableElement(doc, tableIndex);
  const rowElement = requireDefined(
    tableElement.children.filter((child) => isWord(child, "tr"))[row],
    `Missing DOCX row ${row}`,
  );
  const cellElement = requireDefined(
    rowElement.children.filter((child) => isWord(child, "tc"))[column],
    `Missing DOCX cell (${row},${column})`,
  );
  const xml = readDocumentXml(doc);
  return xml.slice(cellElement.start, cellElement.end);
}

function readBody(doc: Document): XmlElement {
  const parsed = parseXml(readDocumentXml(doc));
  return requireDefined(parsed.root.children.find((child) => isWord(child, "body")), "Missing w:body");
}

function readTopLevelTableElement(doc: Document, index: number): XmlElement {
  return requireDefined(
    readBody(doc).children.filter((child) => isWord(child, "tbl"))[index],
    `Missing DOCX table ${index + 1}`,
  );
}

function readDocumentXml(doc: Document): string {
  return UTF8_DECODER.decode(requireDefined(doc.package.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`));
}

function decodeStepText(raw: string): string {
  return raw.replaceAll("\\r", "\r").replaceAll("\\n", "\n");
}

function requireDefined<T>(value: T | undefined, message: string): T {
  expect(value).toBeDefined();
  if (value === undefined) {
    throw new Error(message);
  }
  return value;
}

function isWord(element: XmlElement, localName: string): boolean {
  return element.localName === localName && element.namespaceURI === W_NS;
}
