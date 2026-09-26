import {fixturePath} from "../../scripts/fixture-inputs.ts";
import { expect } from "bun:test";
import { join, resolve } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { Document, W_NS, type Paragraph, type Span } from "../../src/docx/index.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip } from "../../src/opc/zip.ts";
import { attribute, parseXml, type XmlElement } from "../../src/xml/index.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const DOCUMENT_PART = "word/document.xml";
const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });

type AcceptanceState = {
  document?: Document;
  rememberedBytes?: Uint8Array;
  rememberedSpan?: Span;
  rememberedOpaquePart?: { name: string; bytes: Uint8Array };
  lastError?: unknown;
};

export const bindings: StepBinding[] = [
  {
    pattern: /^DOCX create source "([^"]+)" is prepared$/,
    run: async (context, source) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedBytes = undefined;
      state.rememberedSpan = undefined;
      state.rememberedOpaquePart = undefined;
      state.document = source === "new-document"
        ? Document.create()
        : await Document.open(fixturePath(source));
    },
  },
  {
    pattern: /^DOCX create paragraph "([^"]*)" with bold and italic formatting is appended$/,
    run: (context, text) => {
      const state = getState(context);
      state.lastError = undefined;
      document(state).addParagraph(text, { bold: true, italic: true });
    },
  },
  {
    pattern: /^DOCX create styled paragraph "([^"]*)" with style "([^"]+)" is appended$/,
    run: (context, text, style) => {
      const state = getState(context);
      state.lastError = undefined;
      document(state).addParagraph(text, { style });
    },
  },
  {
    pattern: /^DOCX create plain paragraph "([^"]*)" is appended$/,
    run: (context, text) => {
      const state = getState(context);
      state.lastError = undefined;
      document(state).addParagraph(text);
    },
  },
  {
    pattern: /^DOCX create document is saved and reopened$/,
    run: async (context) => {
      const state = getState(context);
      state.lastError = undefined;
      const bytes = await document(state).save();
      state.document = await Document.open(bytes);
    },
  },
  {
    pattern: /^DOCX create package members equal "([^"]+)"$/,
    run: async (context, expectedNames) => {
      const state = getState(context);
      const saved = await document(state).save();
      const names = [...readPackagePartNames(saved)].sort();
      expect(names).toEqual(expectedNames.split(",").sort());
    },
  },
  {
    pattern: /^DOCX create paragraph (\d+) text equals "([^"]*)"$/,
    run: (context, indexText, expected) => {
      const state = getState(context);
      state.lastError = undefined;
      expect(paragraph(state, indexText).text).toBe(expected);
    },
  },
  {
    pattern: /^DOCX create paragraph (\d+) run has bold and italic formatting$/,
    run: (context, indexText) => {
      const run = requireDefined(readParagraphElement(document(getState(context)), Number.parseInt(indexText, 10) - 1)
        .children.find((child) => isWord(child, "r")), "Missing authored DOCX run");
      const runProperties = requireDefined(
        run.children.find((child) => isWord(child, "rPr")),
        "Missing DOCX run properties",
      );
      expect(runProperties.children.some((child) => isWord(child, "b"))).toBeTrue();
      expect(runProperties.children.some((child) => isWord(child, "i"))).toBeTrue();
    },
  },
  {
    pattern: /^DOCX create paragraph (\d+) XML preserves boundary spaces and escapes special characters$/,
    run: (context, indexText) => {
      const xml = rawParagraphXml(document(getState(context)), Number.parseInt(indexText, 10) - 1);
      expect(xml).toContain('<w:t xml:space="preserve">  &lt;Frankenstein &amp; friend&gt;  </w:t>');
    },
  },
  {
    pattern: /^DOCX create paragraph (\d+) is stored before the section properties$/,
    run: (context, indexText) => {
      const paragraphIndex = Number.parseInt(indexText, 10) - 1;
      const body = readBody(document(getState(context)));
      const children = body.children.filter((child) => isWord(child, "p") || isWord(child, "sectPr"));
      expect(children[paragraphIndex]?.localName).toBe("p");
      expect(children[paragraphIndex + 1]?.localName).toBe("sectPr");
    },
  },
  {
    pattern: /^DOCX create paragraph (\d+) uses paragraph style "([^"]+)"$/,
    run: (context, indexText, expectedStyle) => {
      const paragraphElement = readParagraphElement(document(getState(context)), Number.parseInt(indexText, 10) - 1);
      const paragraphProperties = requireDefined(
        paragraphElement.children.find((child) => isWord(child, "pPr")),
        "Missing DOCX paragraph properties",
      );
      const style = requireDefined(
        paragraphProperties.children.find((child) => isWord(child, "pStyle")),
        "Missing DOCX paragraph style",
      );
      expect(attribute(style, "val", W_NS)).toBe(expectedStyle);
    },
  },
  {
    pattern: /^DOCX create paragraph (\d+) exact text "([^"]+)" span is remembered$/,
    run: (context, indexText, query) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedSpan = requireSingleSpan(paragraph(state, indexText).find(query), query);
    },
  },
  {
    pattern: /^DOCX create opaque part "([^"]+)" bytes are remembered$/,
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
    pattern: /^DOCX create current saved bytes are remembered$/,
    run: async (context) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedBytes = await document(state).save();
    },
  },
  {
    pattern: /^DOCX create stale replacement "([^"]+)" is attempted on the remembered span$/,
    run: async (context, replacement) => {
      const state = getState(context);
      state.lastError = undefined;
      try {
        await requireDefined(state.rememberedSpan, "Missing remembered DOCX span").replace(replacement);
      } catch (error) {
        state.lastError = error;
        return;
      }
      throw new Error("Expected stale DOCX span replacement to fail");
    },
  },
  {
    pattern: /^DOCX create append refusal "([^"]+)" is attempted$/,
    run: (context, refusalCase) => {
      const state = getState(context);
      state.lastError = undefined;
      try {
        runRefusalCase(document(state), refusalCase);
      } catch (error) {
        state.lastError = error;
        return;
      }
      throw new Error(`Expected DOCX create refusal for ${refusalCase}`);
    },
  },
  {
    pattern: /^DOCX create refusal code equals "([^"]+)"$/,
    run: (context, expectedCode) => {
      const state = getState(context);
      expect(state.lastError).toBeInstanceOf(OoxmlError);
      expect((state.lastError as OoxmlError).code).toBe(expectedCode);
    },
  },
  {
    pattern: /^DOCX create saved bytes equal the remembered bytes$/,
    run: async (context) => {
      const state = getState(context);
      expect(await document(state).save()).toEqual(requireDefined(state.rememberedBytes, "Missing remembered DOCX bytes"));
    },
  },
  {
    pattern: /^DOCX create opaque part "([^"]+)" bytes are unchanged$/,
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

function paragraph(state: AcceptanceState, indexText: string): Paragraph {
  const index = Number.parseInt(indexText, 10) - 1;
  return requireDefined(document(state).paragraphs[index], `Missing DOCX paragraph ${index + 1}`);
}

function requireSingleSpan(spans: Span[], query: string): Span {
  expect(spans).toHaveLength(1);
  return requireDefined(spans[0], `Expected exactly one DOCX span for ${JSON.stringify(query)}`);
}

function readPackagePartNames(bytes: Uint8Array): Set<string> {
  return new Set(readZip(bytes).keys());
}

function runRefusalCase(doc: Document, refusalCase: string): void {
  switch (refusalCase) {
    case "text-number":
      (doc.addParagraph as (text: unknown) => Paragraph)(7);
      return;
    case "options-string":
      (doc.addParagraph as (text: string, options: unknown) => Paragraph)("text", "bad");
      return;
    case "bold-string":
      (doc.addParagraph as (text: string, options: unknown) => Paragraph)("text", { bold: "yes" });
      return;
    case "style-without-styles-part":
      doc.addParagraph("Heading", { style: "Heading1" });
      return;
    case "unknown-style":
      doc.addParagraph("Heading", { style: "MissingStyle" });
      return;
    default:
      throw new Error(`Unknown DOCX create refusal case ${refusalCase}`);
  }
}

function readBody(doc: Document): XmlElement {
  const parsed = parseXml(readDocumentXml(doc));
  return requireDefined(parsed.root.children.find((child) => isWord(child, "body")), "Missing w:body");
}

function readParagraphElement(doc: Document, index: number): XmlElement {
  return requireDefined(
    readBody(doc).children.filter((child) => isWord(child, "p"))[index],
    `Missing DOCX paragraph ${index + 1}`,
  );
}

function rawParagraphXml(doc: Document, index: number): string {
  const xml = readDocumentXml(doc);
  const paragraphElement = readParagraphElement(doc, index);
  return xml.slice(paragraphElement.start, paragraphElement.end);
}

function readDocumentXml(doc: Document): string {
  return UTF8_DECODER.decode(requireDefined(doc.package.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`));
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
