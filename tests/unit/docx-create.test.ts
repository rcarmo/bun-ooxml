import { describe, expect, test } from "bun:test";
import { join, resolve } from "node:path";

import { bindings as createBindings } from "../acceptance/create-docx.ts";
import { bindings as parentBindings } from "../acceptance/steps.ts";
import {
  executeAcceptance,
  newAcceptanceRunId,
  parseFeature,
  type AcceptanceFeature,
  type AcceptanceInventory,
} from "../../scripts/gherkin.ts";
import { Document, W_NS } from "../../src/docx/index.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip } from "../../src/opc/zip.ts";
import { attribute, parseXml, type XmlElement } from "../../src/xml/index.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const DOCUMENT_PART = "word/document.xml";
const STYLES_FIXTURE = join(
  PROJECT_ROOT,
  "fixtures/python-office-mcp-server/tests/_templates/testdata/word/styles.docx",
);
const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });

describe("Document.create and addParagraph", () => {
  test("creates a minimal package and appends escaped bold italic text before sectPr", async () => {
    const doc = Document.create();

    expect([...doc.package.parts.keys()]).toEqual([
      "[Content_Types].xml",
      "_rels/.rels",
      "word/document.xml",
    ]);
    expect(doc.paragraphs).toHaveLength(0);

    const paragraph = doc.addParagraph("  <Frankenstein & friend>  ", { bold: true, italic: true });
    const saved = await doc.save();
    const reopened = await Document.open(saved);

    expect(paragraph.text).toBe("  <Frankenstein & friend>  ");
    expect(reopened.paragraphs).toHaveLength(1);
    expect(reopened.paragraphs[0]?.text).toBe("  <Frankenstein & friend>  ");
    expect(doc.package.diff()).toEqual({
      added: [],
      changed: [DOCUMENT_PART],
      removed: [],
    });

    const rawXml = readDocumentXml(reopened);
    expect(rawXml).toContain('<w:t xml:space="preserve">  &lt;Frankenstein &amp; friend&gt;  </w:t>');
    const runProperties = requireDefined(
      readParagraphElement(reopened, 0).children.find((child) => isWord(child, "r"))
        ?.children.find((child) => isWord(child, "rPr")),
      "Missing run properties",
    );
    expect(runProperties.children.some((child) => isWord(child, "b"))).toBeTrue();
    expect(runProperties.children.some((child) => isWord(child, "i"))).toBeTrue();

    const bodyChildren = readBody(reopened).children.filter((child) => isWord(child, "p") || isWord(child, "sectPr"));
    expect(bodyChildren.map((child) => child.localName)).toEqual(["p", "sectPr"]);
  });

  test("permits empty text and keeps the authored paragraph before sectPr after reopen", async () => {
    const doc = Document.create();
    const paragraph = doc.addParagraph();
    const reopened = await Document.open(await doc.save());

    expect(paragraph.text).toBe("");
    expect(reopened.paragraphs).toHaveLength(1);
    expect(reopened.paragraphs[0]?.text).toBe("");
    expect(readBody(reopened).children.filter((child) => isWord(child, "p") || isWord(child, "sectPr"))
      .map((child) => child.localName)).toEqual(["p", "sectPr"]);
  });

  test("validates known paragraph styles and preserves opaque parts on append", async () => {
    const doc = await Document.open(STYLES_FIXTURE);
    const originalCore = requireDefined(doc.package.get("docProps/core.xml"), "Missing docProps/core.xml");
    const originalStyles = requireDefined(doc.package.get("word/styles.xml"), "Missing word/styles.xml");

    const paragraph = doc.addParagraph("Created heading", { style: "Heading1" });
    const reopened = await Document.open(await doc.save());

    expect(paragraph.text).toBe("Created heading");
    expect(reopened.paragraphs).toHaveLength(5);
    expect(reopened.paragraphs[4]?.text).toBe("Created heading");
    expect(attribute(
      requireDefined(
        readParagraphElement(reopened, 4).children.find((child) => isWord(child, "pPr"))
          ?.children.find((child) => isWord(child, "pStyle")),
        "Missing paragraph style element",
      ),
      "val",
      W_NS,
    )).toBe("Heading1");
    expect(doc.package.get("docProps/core.xml")).toEqual(originalCore);
    expect(doc.package.get("word/styles.xml")).toEqual(originalStyles);
    expect(doc.package.diff()).toEqual({
      added: [],
      changed: [DOCUMENT_PART],
      removed: [],
    });
  });

  test("refuses invalid arguments and unsupported or unknown styles atomically", async () => {
    const newDocument = Document.create();
    const newBaseline = await newDocument.save();

    expect(() => (newDocument.addParagraph as (text: unknown) => unknown)(7)).toThrow(
      expect.objectContaining({ code: "docx-invalid-argument" }),
    );
    expect(() => (newDocument.addParagraph as (text: string, options: unknown) => unknown)("text", "bad")).toThrow(
      expect.objectContaining({ code: "docx-invalid-argument" }),
    );
    expect(() => (newDocument.addParagraph as (text: string, options: unknown) => unknown)("text", { bold: "yes" })).toThrow(
      expect.objectContaining({ code: "docx-invalid-argument" }),
    );
    expect(() => newDocument.addParagraph("Heading", { style: "Heading1" })).toThrow(
      expect.objectContaining({ code: "docx-style-unsupported" }),
    );
    expect(await newDocument.save()).toEqual(newBaseline);

    const styledDocument = await Document.open(STYLES_FIXTURE);
    const styledBaseline = await styledDocument.save();
    expect(() => styledDocument.addParagraph("Heading", { style: "MissingStyle" })).toThrow(
      expect.objectContaining({ code: "docx-style-missing" }),
    );
    expect(await styledDocument.save()).toEqual(styledBaseline);
  });

  test("invalidates earlier spans after append without disturbing preserved parts", async () => {
    const doc = await Document.open(STYLES_FIXTURE);
    const stale = requireSingleSpan(requireDefined(doc.paragraphs[1], "Missing paragraph 2").find("Genevese"), "Genevese");
    const originalCore = requireDefined(doc.package.get("docProps/core.xml"), "Missing docProps/core.xml");

    doc.addParagraph("Appendix line");
    const bytesAfterAppend = await doc.save();

    let refusal: unknown;
    try {
      await stale.replace("Citizen");
    } catch (error) {
      refusal = error;
    }

    expect(refusal).toBeInstanceOf(OoxmlError);
    expect((refusal as OoxmlError).code).toBe("docx-stale-span");
    expect(await doc.save()).toEqual(bytesAfterAppend);
    expect(doc.package.get("docProps/core.xml")).toEqual(originalCore);
  });

  test("executes the DOCX create feature with dedicated bindings and optional parent aggregation", async () => {
    const featurePath = "features/docx/create.feature";
    const feature = parseFeature(featurePath, await Bun.file(join(PROJECT_ROOT, featurePath)).text());
    const inventory = inventoryFor(feature);

    for (const bindings of [createBindings, parentBindings]) {
      const execution = await executeAcceptance(inventory, bindings, newAcceptanceRunId());
      expect(execution.failures).toEqual([]);
      expect(execution.counts.features).toEqual({ passed: 1, failed: 0, planned: 0, total: 1 });
      expect(execution.counts.scenarios).toEqual({ passed: 4, failed: 0, planned: 0, total: 4 });
      expect(execution.counts.cases).toEqual({ passed: 8, failed: 0, planned: 0, total: 8 });
      expect(execution.counts.steps.failed).toBe(0);
      expect(execution.counts.steps.undefined).toBe(0);
      expect(execution.counts.steps.ambiguous).toBe(0);
    }
  }, 300_000);
});

function inventoryFor(feature: AcceptanceFeature): AcceptanceInventory {
  const cases = feature.scenarios.flatMap((scenario) => scenario.cases);
  const steps = cases.flatMap((acceptanceCase) => acceptanceCase.steps).length;
  const count = (implemented: number) => ({ implemented, planned: 0, total: implemented });
  return {
    root: PROJECT_ROOT,
    features: [feature],
    counts: {
      features: count(1),
      scenarios: count(feature.scenarios.length),
      cases: count(cases.length),
      steps: count(steps),
    },
  };
}

function readBody(doc: Document): XmlElement {
  const parsed = parseXml(readDocumentXml(doc));
  return requireDefined(parsed.root.children.find((child) => isWord(child, "body")), "Missing w:body");
}

function readParagraphElement(doc: Document, index: number): XmlElement {
  return requireDefined(
    readBody(doc).children.filter((child) => isWord(child, "p"))[index],
    `Missing paragraph ${index + 1}`,
  );
}

function readDocumentXml(doc: Document): string {
  return UTF8_DECODER.decode(requireDefined(doc.package.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`));
}

function requireSingleSpan(spans: import("../../src/docx/index.ts").Span[], query: string) {
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

function isWord(element: XmlElement, localName: string): boolean {
  return element.localName === localName && element.namespaceURI === W_NS;
}
