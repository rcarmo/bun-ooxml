import {fixturePath,fixturePaths,F} from "../../scripts/fixture-inputs.ts";
import { expect } from "bun:test";
import { readFileSync } from "node:fs";
import { join, resolve } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { Document, W_NS, type Paragraph, type Span } from "../../src/docx/index.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { applyEdits, parseXml, type XmlElement } from "../../src/xml/index.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const DOCUMENT_PART = "word/document.xml";
const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });
const UTF8_ENCODER = new TextEncoder();

type AcceptanceState = {
  document?: Document;
  rememberedBytes?: Uint8Array;
  rememberedSpan?: Span;
  lastError?: unknown;
};

export const bindings: StepBinding[] = [
  {
    pattern: /^DOCX slice fixture "([^"]+)" is opened$/,
    run: async (context, fixture) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedBytes = undefined;
      state.rememberedSpan = undefined;
      state.document = await Document.open(resolveFixture(fixture));
    },
  },
  {
    pattern: /^DOCX slice paragraph (\d+) exact text "([^"]+)" is replaced with "([^"]+)"$/,
    run: async (context, indexText, query, replacement) => {
      const state = getState(context);
      state.lastError = undefined;
      const span = requireSingleSpan(paragraph(state, indexText).find(query), query);
      await span.replace(replacement);
    },
  },
  {
    pattern: /^DOCX slice document is saved and reopened$/,
    run: async (context) => {
      const state = getState(context);
      state.lastError = undefined;
      state.document = await Document.open(await document(state).save());
    },
  },
  {
    pattern: /^DOCX slice paragraph (\d+) text equals "([^"]+)"$/,
    run: (context, indexText, expected) => {
      const state = getState(context);
      state.lastError = undefined;
      expect(paragraph(state, indexText).text).toBe(expected);
    },
  },
  {
    pattern: /^DOCX slice paragraph (\d+) run formatting around the replacement is preserved$/,
    run: (context, indexText) => {
      const state = getState(context);
      state.lastError = undefined;
      const runs = paragraphRunSnapshot(document(state), Number.parseInt(indexText, 10) - 1);
      expect(runs).toEqual([
        { text: "Bold Tone", styles: ["b"] },
        { text: " and ", styles: [] },
        { text: "tilted text", styles: ["i"] },
        { text: " and ", styles: [] },
        { text: "underlined text", styles: ["u"] },
        { text: " and ", styles: [] },
        { text: "colored text", styles: ["color"] },
        { text: " and ", styles: [] },
        { text: "large text", styles: ["sz"] },
      ]);
    },
  },
  {
    pattern: /^DOCX slice paragraph (\d+) text nodes preserve boundary whitespace$/,
    run: (context, indexText) => {
      const state = getState(context);
      state.lastError = undefined;
      const texts = paragraphTextSnapshot(document(state), Number.parseInt(indexText, 10) - 1);
      expect(texts).toEqual([
        { text: "A", preserve: undefined },
        { text: "lpha", preserve: undefined },
        { text: " Beta", preserve: "preserve" },
      ]);
    },
  },
  {
    pattern: /^DOCX slice paragraph (\d+) span "([^"]+)" is remembered$/,
    run: (context, indexText, query) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedSpan = requireSingleSpan(paragraph(state, indexText).find(query), query);
    },
  },
  {
    pattern: /^DOCX slice current saved bytes are remembered$/,
    run: async (context) => {
      const state = getState(context);
      state.lastError = undefined;
      state.rememberedBytes = await document(state).save();
    },
  },
  {
    pattern: /^DOCX slice stale replacement "([^"]+)" is attempted on the remembered span$/,
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
    pattern: /^DOCX slice refusal code equals "([^"]+)"$/,
    run: (context, expectedCode) => {
      const state = getState(context);
      expect(state.lastError).toBeInstanceOf(OoxmlError);
      expect((state.lastError as OoxmlError).code).toBe(expectedCode);
    },
  },
  {
    pattern: /^DOCX slice saved bytes equal the remembered bytes$/,
    run: async (context) => {
      const state = getState(context);
      const saved = await document(state).save();
      if (!state.rememberedBytes) throw new Error("Missing remembered DOCX bytes");
      expect(saved).toEqual(state.rememberedBytes);
    },
  },
  {
    pattern: /^DOCX slice paragraph (\d+) exact search for "([^"]+)" is attempted$/,
    run: (context, indexText, query) => {
      const state = getState(context);
      state.lastError = undefined;
      try {
        paragraph(state, indexText).find(query);
      } catch (error) {
        state.lastError = error;
        return;
      }
      throw new Error("Expected DOCX exact search refusal");
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

function requireDefined<T>(value: T | undefined, message: string): T {
  expect(value).toBeDefined();
  if (value === undefined) {
    throw new Error(message);
  }
  return value;
}

function resolveFixture(fixture: string): string | Uint8Array {
  if (fixture === "synthetic-field") {
    return buildSyntheticFieldFixture();
  }
  if (fixture === "synthetic-whitespace") {
    return buildSyntheticWhitespaceFixture();
  }
  return fixturePath(fixture);
}

function buildSyntheticWhitespaceFixture(): Uint8Array {
  return replaceFirstParagraph(
    [
      '<w:r><w:t>A</w:t></w:r>',
      '<w:r><w:t>lpha</w:t></w:r>',
      '<w:r><w:t>Beta</w:t></w:r>',
    ].join(""),
  );
}

function buildSyntheticFieldFixture(): Uint8Array {
  return replaceFirstParagraph(
    [
      '<w:r><w:t xml:space="preserve">Before </w:t></w:r>',
      '<w:fldSimple w:instr="DATE"><w:r><w:t>2026-01-01</w:t></w:r></w:fldSimple>',
      '<w:r><w:t xml:space="preserve"> After</w:t></w:r>',
    ].join(""),
  );
}

function replaceFirstParagraph(innerXml: string): Uint8Array {
  const source = new Uint8Array(readFileSync(fixturePath(F.officeWord.singleParagraph)));
  const parts = readZip(source);
  const xml = UTF8_DECODER.decode(requireDefined(parts.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`));
  const parsed = parseXml(xml);
  const paragraph = findFirstParagraph(parsed.root);
  const updated = applyEdits(xml, [
    {
      start: paragraph.openEnd,
      end: paragraph.closeStart,
      value: innerXml,
    },
  ]);
  parts.set(DOCUMENT_PART, UTF8_ENCODER.encode(updated));
  return writeZip(parts);
}

function findFirstParagraph(root: XmlElement): XmlElement {
  const body = requireDefined(
    root.children.find((child) => child.localName === "body" && child.namespaceURI === W_NS),
    "Missing w:body",
  );
  return requireDefined(
    body.children.find((child) => child.localName === "p" && child.namespaceURI === W_NS),
    "Missing first w:p",
  );
}

function paragraphRunSnapshot(doc: Document, index: number): Array<{ text: string; styles: string[] }> {
  const paragraph = readParagraphElement(doc, index);
  return paragraph.children
    .filter((child) => child.localName === "r" && child.namespaceURI === W_NS)
    .map((run) => ({
      text: run.children
        .filter((child) => child.localName === "t" && child.namespaceURI === W_NS)
        .map((child) => child.text)
        .join(""),
      styles:
        run.children
          .find((child) => child.localName === "rPr" && child.namespaceURI === W_NS)
          ?.children.filter((child) => child.namespaceURI === W_NS)
          .map((child) => child.localName) ?? [],
    }));
}

function paragraphTextSnapshot(
  doc: Document,
  index: number,
): Array<{ text: string; preserve: string | undefined }> {
  const paragraph = readParagraphElement(doc, index);
  return paragraph.children
    .filter((child) => child.localName === "r" && child.namespaceURI === W_NS)
    .flatMap((run) =>
      run.children
        .filter((child) => child.localName === "t" && child.namespaceURI === W_NS)
        .map((child) => ({ text: child.text, preserve: child.attributes["xml:space"] })),
    );
}

function readParagraphElement(doc: Document, index: number): XmlElement {
  const xml = UTF8_DECODER.decode(
    requireDefined(doc.package.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`),
  );
  const parsed = parseXml(xml);
  const body = requireDefined(
    parsed.root.children.find((child) => child.localName === "body" && child.namespaceURI === W_NS),
    "Missing w:body",
  );
  return requireDefined(
    body.children.filter((child) => child.localName === "p" && child.namespaceURI === W_NS)[index],
    `Missing paragraph ${index + 1}`,
  );
}
