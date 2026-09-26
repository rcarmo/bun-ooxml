import {fixturePath,fixturePaths,F} from "../../scripts/fixture-inputs.ts";
import { afterEach, describe, expect, test } from "bun:test";
import {textboxDocument} from '../fixtures/native-edge-cases.ts';
import { readFileSync, readdirSync } from "node:fs";
import { copyFile, mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { bindings as docxBindings } from "../acceptance/docx.ts";
import { Document, W_NS, type Paragraph, type Span } from "../../src/docx/index.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { applyEdits, parseXml, type XmlElement } from "../../src/xml/index.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const DOCUMENT_PART = "word/document.xml";
const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });
const UTF8_ENCODER = new TextEncoder();
const tempRoots: string[] = [];

afterEach(async () => {
  await Promise.all(tempRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("docx Document", () => {
  test("opens from a path or bytes and save() never rewrites the opened source path implicitly", async () => {
    const tempPath = await copyFixture(fixturePath(F.officeWord.singleParagraph));
    const original = new Uint8Array(readFileSync(tempPath));

    const fromPath = await Document.open(tempPath);
    const fromBytes = await Document.open(original);

    const expected =
      "I am by birth a Genevese, and my family is one of the most distinguished of that republic.";

    expect(fromPath.package.get(DOCUMENT_PART)).toBeDefined();
    expect(fromPath.paragraphs).toHaveLength(1);
    expect(requireParagraph(fromPath, 0).text).toBe(expected);
    expect(requireParagraph(fromBytes, 0).text).toBe(expected);
    expect(await fromPath.save()).toEqual(original);
    expect(new Uint8Array(readFileSync(tempPath))).toEqual(original);
    expect(await fromBytes.save()).toEqual(original);

    const explicitPath = join(dirname(tempPath), "saved.docx");
    expect(await fromBytes.save(explicitPath)).toEqual(original);
    expect(new Uint8Array(readFileSync(explicitPath))).toEqual(original);
  });

  test("preserves exact bytes for unmodified word fixtures from the go-ooxml and python corpora across save and reopen", async () => {
    for (const path of wordFixtures()) {
      const original = new Uint8Array(readFileSync(path));
      const doc = await Document.open(original);
      const saved = await doc.save();
      expect(saved).toEqual(original);

      const reopened = await Document.open(saved);
      expect(await reopened.save()).toEqual(original);
    }
  });

  test("replaces exact cross-run text without reconstructing unaffected runs or formatting", async () => {
    const tempPath = await copyFixture(fixturePath(F.officeWord.formattedText));
    const original = new Uint8Array(readFileSync(tempPath));
    const doc = await Document.open(tempPath);

    const first = requireParagraph(doc, 0);
    const span = requireSingleSpan(first.find("text and italic text"), "text and italic text");

    await span.replace("Tone and tilted text");

    const saved = await doc.save();
    expect(new Uint8Array(readFileSync(tempPath))).toEqual(original);

    const reopened = await Document.open(saved);
    expect(requireParagraph(reopened, 0).text).toBe(
      "Bold Tone and tilted text and underlined text and colored text and large text",
    );
    expect(requireParagraph(reopened, 1).text).toBe(
      "I saw the dull yellow eye of the creature open; it breathed hard, and a convulsive motion agitated its limbs.",
    );

    expect(paragraphRunSnapshot(reopened, 0)).toEqual([
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
  });

  test("preserves boundary whitespace with xml:space on touched text nodes", async () => {
    const doc = await Document.open(buildWhitespaceFixture());
    const span = requireSingleSpan(requireParagraph(doc, 0).find("phaBe"), "phaBe");

    await span.replace("pha Be");

    const reopened = await Document.open(await doc.save());
    expect(requireParagraph(reopened, 0).text).toBe("Alpha Beta");
    expect(paragraphTextSnapshot(reopened, 0)).toEqual([
      { text: "A", preserve: undefined },
      { text: "lpha", preserve: undefined },
      { text: " Beta", preserve: "preserve" },
    ]);
  });

  test("includes table-cell paragraphs in document order and replaces them without reconstructing unrelated table XML", async () => {
    const original = new Uint8Array(readFileSync(fixturePath(F.officeWord.simpleTable)));
    const originalXml = UTF8_DECODER.decode(
      requireDefined(readZip(original).get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`),
    );
    const doc = await Document.open(original);

    expect(doc.paragraphs.map((paragraph) => paragraph.text)).toEqual([
      "Research Materials Inventory",
      "Item",
      "Quantity",
      "Source",
      "Galvanic battery",
      "3",
      "Germany",
      "Chemical reagents",
      "12 vials",
      "Switzerland",
      "Anatomical charts",
      "7",
      "University library",
      "",
    ]);

    const span = requireSingleSpan(doc.find("Galvanic battery"), "Galvanic battery");
    await span.replace("Voltaic battery");

    const saved = await doc.save();
    const savedXml = UTF8_DECODER.decode(
      requireDefined(readZip(saved).get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`),
    );
    expect(savedXml).toBe(originalXml.replace(">Galvanic battery<", ">Voltaic battery<"));

    const reopened = await Document.open(saved);
    expect(reopened.paragraphs.map((paragraph) => paragraph.text)).toEqual([
      "Research Materials Inventory",
      "Item",
      "Quantity",
      "Source",
      "Voltaic battery",
      "3",
      "Germany",
      "Chemical reagents",
      "12 vials",
      "Switzerland",
      "Anatomical charts",
      "7",
      "University library",
      "",
    ]);
  });

  test("does not descend into text boxes and refuses the blind region when no supported match exists", async () => {
    const doc = await Document.open(await textboxDocument());

    expect(doc.find("Outside the text box.")).toHaveLength(1);
    expect(() => doc.find("Text living inside the text box.")).toThrow(
      expect.objectContaining({ code: "docx-unsupported-topology" }),
    );
  });

  test("detects stale spans before mutation and leaves the package unchanged", async () => {
    const doc = await Document.open(fixturePath(F.officeWord.formattedText));

    const first = requireParagraph(doc, 0);
    const stale = requireSingleSpan(first.find("text and italic text"), "text and italic text");
    const colored = requireSingleSpan(first.find("colored text"), "colored text");
    await colored.replace("scarlet text");
    const baseline = await doc.save();

    let refusal: unknown;
    try {
      await stale.replace("Tone and tilted text");
    } catch (error) {
      refusal = error;
    }

    expect(refusal).toBeInstanceOf(OoxmlError);
    expect((refusal as OoxmlError).code).toBe("docx-stale-span");
    expect(await doc.save()).toEqual(baseline);
  });

  test("keeps repeated exact occurrences correct across replacements", async () => {
    const doc = await Document.open(buildRepeatedFixture());
    const paragraph = requireParagraph(doc, 0);

    const matches = paragraph.find("Alpha Beta");
    expect(matches).toHaveLength(2);

    await requireDefined(matches[0], "Missing first repeated DOCX span").replace("Gamma Delta");

    const refreshed = requireParagraph(doc, 0).find("Alpha Beta");
    expect(refreshed).toHaveLength(1);
    await requireDefined(refreshed[0], "Missing second repeated DOCX span").replace("Gamma Delta");

    const reopened = await Document.open(await doc.save());
    expect(requireParagraph(reopened, 0).text).toBe("Gamma Delta Gamma Delta");
  });

  test("refuses fields, tracked revisions and content controls with a stable code", async () => {
    const fieldDoc = await Document.open(buildFieldFixture());
    const trackedDoc = await Document.open(fixturePath(F.officeWord.trackChanges));
    const controlsDoc = await Document.open(fixturePath(F.officeWord.sdtContentControls));

    expect(() => requireParagraph(fieldDoc, 0).find("2026-01-01")).toThrow(
      expect.objectContaining({ code: "docx-unsupported-topology" }),
    );
    expect(() => requireParagraph(trackedDoc, 0).find("amazing")).toThrow(
      expect.objectContaining({ code: "docx-unsupported-topology" }),
    );
    expect(() => requireParagraph(controlsDoc, 1).find("[Enter Title Here]")).toThrow(
      expect.objectContaining({ code: "docx-unsupported-topology" }),
    );
  });

  test("passes the DOCX acceptance feature with exported bindings", async () => {
    const root = await makeAcceptanceRoot();
    const report = await runAcceptance(docxBindings, { root });

    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.inventory.scenarios.implemented).toBe(5);
    expect(report.execution.cases.passed).toBe(7);
    expect(report.execution.steps.failed).toBe(0);
    expect(report.failures).toEqual([]);
  });
});

function wordFixtures(): string[] { return fixturePaths([...Object.values(F.goWord),...Object.values(F.officeWord)].filter(id=>id!==F.goWord.default&&id!==F.officeWord.default)); }

function paragraphRunSnapshot(doc: Document, index: number): Array<{ text: string; styles: string[] }> {
  const paragraph = readParagraphElement(doc, index);
  return paragraph.children
    .filter((child) => isWord(child, "r"))
    .map((run) => ({
      text: run.children.filter((child) => isWord(child, "t")).map((child) => child.text).join(""),
      styles:
        run.children
          .find((child) => isWord(child, "rPr"))
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
    .filter((child) => isWord(child, "r"))
    .flatMap((run) =>
      run.children
        .filter((child) => isWord(child, "t"))
        .map((child) => ({ text: child.text, preserve: child.attributes["xml:space"] })),
    );
}

function buildRepeatedFixture(): Uint8Array {
  return replaceFirstParagraph(
    [
      '<w:r><w:t xml:space="preserve">Alpha </w:t></w:r>',
      '<w:r><w:t>Beta</w:t></w:r>',
      '<w:r><w:t xml:space="preserve"> Alpha </w:t></w:r>',
      '<w:r><w:t>Beta</w:t></w:r>',
    ].join(""),
  );
}

function buildWhitespaceFixture(): Uint8Array {
  return replaceFirstParagraph(
    [
      '<w:r><w:t>A</w:t></w:r>',
      '<w:r><w:t>lpha</w:t></w:r>',
      '<w:r><w:t>Beta</w:t></w:r>',
    ].join(""),
  );
}

function buildFieldFixture(): Uint8Array {
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
  const xmlBytes = requireDefined(parts.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`);
  const xml = UTF8_DECODER.decode(xmlBytes);
  const parsed = parseXml(xml);
  const paragraph = findFirstParagraph(parsed.root);
  const updated = applyEdits(xml, [
    { start: paragraph.openEnd, end: paragraph.closeStart, value: innerXml },
  ]);
  parts.set(DOCUMENT_PART, UTF8_ENCODER.encode(updated));
  return writeZip(parts);
}

function findFirstParagraph(root: XmlElement): XmlElement {
  const body = requireDefined(root.children.find((child) => isWord(child, "body")), "Missing w:body");
  return requireDefined(body.children.find((child) => isWord(child, "p")), "Missing first w:p");
}

function readParagraphElement(doc: Document, index: number): XmlElement {
  const xmlBytes = requireDefined(doc.package.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`);
  const xml = UTF8_DECODER.decode(xmlBytes);
  const parsed = parseXml(xml);
  const body = requireDefined(
    parsed.root.children.find((child) => isWord(child, "body")),
    "Missing w:body",
  );
  return requireDefined(
    body.children.filter((child) => isWord(child, "p"))[index],
    `Missing paragraph ${index + 1}`,
  );
}

function requireParagraph(doc: Document, index: number): Paragraph {
  return requireDefined(doc.paragraphs[index], `Missing DOCX paragraph ${index + 1}`);
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

function isWord(element: XmlElement, localName: string): boolean {
  return element.localName === localName && element.namespaceURI === W_NS;
}

async function makeAcceptanceRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-docx-acceptance-"));
  tempRoots.push(root);
  const featureText = await Bun.file(join(PROJECT_ROOT, "features/docx/text.feature")).text();
  const path = join(root, "features", "docx", "text.feature");
  await mkdir(dirname(path), { recursive: true });
  await Bun.write(path, featureText);
  return root;
}

async function copyFixture(sourcePath: string): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-docx-fixture-"));
  tempRoots.push(root);
  const target = join(root, basename(sourcePath));
  await copyFile(sourcePath, target);
  return target;
}
