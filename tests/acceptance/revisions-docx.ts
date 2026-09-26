import { expect } from "bun:test";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { Document } from "../../src/docx/index.ts";
import { OoxmlError } from "../../src/errors.ts";
import { OpcPackage, addPart, addRelationship } from "../../src/opc/index.ts";
import { inspectRevisions, resolveRevisions } from "../../src/docx/revisions.ts";
import { parseXml } from "../../src/xml/index.ts";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const FOOTNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const ENDNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const SETTINGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

type RevisionView = "current" | "original" | "all";

type RevisionAcceptanceState = {
  pkg?: OpcPackage;
  changedParts?: string[];
  rememberedBytes?: Uint8Array;
  lastError?: unknown;
};

export const bindings: StepBinding[] = [
  {
    pattern: /^the synthetic DOCX revisions fixture "([^"]+)" is opened$/,
    run: async (context, name) => {
      const current = state(context);
      current.changedParts = undefined;
      current.rememberedBytes = undefined;
      current.lastError = undefined;
      current.pkg = await openRevisionFixture(name);
    },
  },
  {
    pattern: /^the current DOCX revision bytes are remembered$/,
    run: (context) => {
      state(context).rememberedBytes = pkg(context).toBytes();
    },
  },
  {
    pattern: /^the DOCX revision story texts for view "([^"]+)" are:?$/,
    run: async (context, viewText) => {
      const view = asView(viewText);
      const expected = readStoryTextRows(context);
      const actual = await Promise.all(expected.map(async ({ part }) => ({
        part,
        text: await readRevisionStoryText(pkg(context), part, view),
      })));
      expect(actual).toEqual(expected);
    },
  },
  {
    pattern: /^all DOCX revisions are (accepted|rejected)$/,
    run: (context, action) => {
      const current = state(context);
      current.lastError = undefined;
      try {
        current.changedParts = resolveRevisions(pkg(context), action === "accepted" ? "accept" : "reject").changedParts;
      } catch (error) {
        current.lastError = error;
      }
    },
  },
  {
    pattern: /^DOCX revisions are accepted for parts$/,
    run: (context) => {
      const current = state(context);
      current.lastError = undefined;
      try {
        current.changedParts = resolveRevisions(pkg(context), "accept", { parts: readPartRows(context) }).changedParts;
      } catch (error) {
        current.lastError = error;
      }
    },
  },
  {
    pattern: /^the DOCX revision package is saved and reopened$/,
    run: async (context) => {
      state(context).pkg = await OpcPackage.open(pkg(context).toBytes());
    },
  },
  {
    pattern: /^the DOCX changed parts are$/,
    run: (context) => {
      expect(state(context).changedParts ?? []).toEqual(readPartRows(context));
    },
  },
  {
    pattern: /^the DOCX revision count is (\d+)$/,
    run: (context, countText) => {
      expect(inspectRevisions(pkg(context)).revisions).toHaveLength(Number.parseInt(countText, 10));
    },
  },
  {
    pattern: /^the DOCX revision refusal code equals "([^"]+)"$/,
    run: (context, expectedCode) => {
      const error = state(context).lastError;
      expect(error).toBeInstanceOf(OoxmlError);
      expect((error as OoxmlError).code).toBe(expectedCode);
    },
  },
  {
    pattern: /^the DOCX revision refusal mentions "([^"]+)"$/,
    run: (context, fragment) => {
      const error = state(context).lastError;
      expect(error).toBeInstanceOf(OoxmlError);
      expect((error as OoxmlError).message).toContain(fragment);
    },
  },
  {
    pattern: /^the DOCX revision saved bytes equal the remembered bytes$/,
    run: (context) => {
      expect(pkg(context).toBytes()).toEqual(required(state(context).rememberedBytes, "missing remembered DOCX revision bytes"));
    },
  },
];

export async function openRevisionFixture(name: string): Promise<OpcPackage> {
  if (name === "empty-deletion") {
    const pkg = await blankPackage();
    pkg.set(pkg.mainPart(), `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:del w:id="1"><w:r><w:delText/></w:r></w:del></w:p></w:body></w:document>`);
    return pkg;
  }
  if (name === "all-stories") {
    return createAllStoriesFixture();
  }
  if (name === "rollback-unsupported") {
    return createRollbackUnsupportedFixture();
  }
  if (name === "unsafe-lift") {
    return createUnsafeLiftFixture();
  }
  if (name === "protected") {
    return createProtectedFixture();
  }
  throw new Error(`unknown DOCX revisions fixture ${name}`);
}

export async function readRevisionStoryText(
  source: OpcPackage,
  part: string,
  view: RevisionView,
): Promise<string> {
  if (view === "all") {
    return extractStoryText(source.text(part), true);
  }

  const clone = await OpcPackage.open(source.toBytes());
  resolveRevisions(clone, view === "current" ? "accept" : "reject", { parts: [part] });
  return extractStoryText(clone.text(part), false);
}

async function createAllStoriesFixture(): Promise<OpcPackage> {
  const pkg = await blankPackage();
  const main = pkg.mainPart();

  pkg.set(main, [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:document xmlns:w="${W_NS}">`,
    "  <w:body>",
    `    ${revisionParagraph("1", "2", "Body historic.", "Body modernized.")}`,
    "    <w:sectPr/>",
    "  </w:body>",
    "</w:document>",
  ].join("\n"));

  addPart(pkg, "word/header1.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:hdr xmlns:w="${W_NS}">`,
    `  ${revisionParagraph("3", "4", "Header historic.", "Header modernized.")}`,
    "</w:hdr>",
  ].join("\n"), HEADER_CONTENT_TYPE);

  addPart(pkg, "word/footer1.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:ftr xmlns:w="${W_NS}">`,
    `  ${revisionParagraph("5", "6", "Footer historic.", "Footer modernized.")}`,
    "</w:ftr>",
  ].join("\n"), FOOTER_CONTENT_TYPE);

  addPart(pkg, "word/footer2.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:ftr xmlns:w="${W_NS}">`,
    `  ${namespacedRevisionParagraph("7", "8", "Nested footer historic.", "Nested footer modernized.")}`,
    "</w:ftr>",
  ].join("\n"), FOOTER_CONTENT_TYPE);

  addPart(pkg, "word/footnotes.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:footnotes xmlns:w="${W_NS}">`,
    `  <w:footnote w:id="1">${revisionParagraph("9", "10", "Footnote historic.", "Footnote modernized.")}</w:footnote>`,
    "</w:footnotes>",
  ].join("\n"), FOOTNOTES_CONTENT_TYPE);

  addPart(pkg, "word/endnotes.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:endnotes xmlns:w="${W_NS}">`,
    `  <w:endnote w:id="1">${revisionParagraph("11", "12", "Endnote historic.", "Endnote modernized.")}</w:endnote>`,
    "</w:endnotes>",
  ].join("\n"), ENDNOTES_CONTENT_TYPE);

  addPart(pkg, "word/comments.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:comments xmlns:w="${W_NS}">`,
    `  <w:comment w:id="0" w:author="A">${revisionParagraph("13", "14", "Comment historic.", "Comment modernized.")}</w:comment>`,
    "</w:comments>",
  ].join("\n"), COMMENTS_CONTENT_TYPE);

  addRelationship(pkg, main, `${OFFICE_REL}header`, "header1.xml");
  addRelationship(pkg, main, `${OFFICE_REL}footer`, "footer1.xml");
  addRelationship(pkg, main, `${OFFICE_REL}footnotes`, "footnotes.xml");
  addRelationship(pkg, main, `${OFFICE_REL}endnotes`, "endnotes.xml");
  addRelationship(pkg, main, `${OFFICE_REL}comments`, "comments.xml");
  addRelationship(pkg, "word/header1.xml", `${OFFICE_REL}footer`, "footer2.xml");

  return pkg;
}

async function createRollbackUnsupportedFixture(): Promise<OpcPackage> {
  const pkg = await blankPackage();
  const main = pkg.mainPart();

  pkg.set(main, [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:document xmlns:w="${W_NS}">`,
    "  <w:body>",
    `    ${revisionParagraph("1", "2", "Body old.", "Body safe.")}`,
    "    <w:sectPr/>",
    "  </w:body>",
    "</w:document>",
  ].join("\n"));

  addPart(pkg, "word/header1.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>Header link.</w:t></w:r></w:p></w:hdr>`,
  ].join("\n"), HEADER_CONTENT_TYPE);

  addPart(pkg, "word/footer2.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:ftr xmlns:w="${W_NS}">`,
    "  <w:p>",
    "    <w:r>",
    "      <w:rPr><w:rPrChange w:id=\"91\" w:author=\"A\" w:date=\"2026-06-01T09:30:00Z\"><w:rPr/></w:rPrChange></w:rPr>",
    "      <w:t>Unsafe footer.</w:t>",
    "    </w:r>",
    "  </w:p>",
    "</w:ftr>",
  ].join("\n"), FOOTER_CONTENT_TYPE);

  addRelationship(pkg, main, `${OFFICE_REL}header`, "header1.xml");
  addRelationship(pkg, "word/header1.xml", `${OFFICE_REL}footer`, "footer2.xml");

  return pkg;
}

async function createUnsafeLiftFixture(): Promise<OpcPackage> {
  const pkg = await blankPackage();
  const main = pkg.mainPart();

  pkg.set(main, [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:document xmlns:w="${W_NS}">`,
    "  <w:body>",
    "    <w:p>",
    `      <w:del xmlns:w14="${W14_NS}" w:id="1" w:author="A" w:date="2026-06-01T09:30:00Z">`,
    `        <w:r xmlns:w14="urn:test:conflict"><w:rPr><w14:glow/></w:rPr><w:delText>Conflict text.</w:delText></w:r>`,
    "      </w:del>",
    "    </w:p>",
    "    <w:sectPr/>",
    "  </w:body>",
    "</w:document>",
  ].join("\n"));

  return pkg;
}

async function createProtectedFixture(): Promise<OpcPackage> {
  const pkg = await blankPackage();
  const main = pkg.mainPart();

  pkg.set(main, [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:document xmlns:w="${W_NS}">`,
    "  <w:body>",
    `    ${revisionParagraph("1", "2", "Locked old.", "Locked new.")}`,
    "    <w:sectPr/>",
    "  </w:body>",
    "</w:document>",
  ].join("\n"));

  addPart(pkg, "word/settings.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:settings xmlns:w="${W_NS}">`,
    '  <w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>',
    "</w:settings>",
  ].join("\n"), SETTINGS_CONTENT_TYPE);
  addRelationship(pkg, main, `${OFFICE_REL}settings`, "settings.xml");

  return pkg;
}

async function blankPackage(): Promise<OpcPackage> {
  return OpcPackage.open(await Document.create().save());
}

function revisionParagraph(deleteId: string, insertId: string, original: string, current: string): string {
  return [
    "<w:p>",
    `  <w:del w:id="${deleteId}" w:author="A" w:date="2026-06-01T09:30:00Z"><w:r><w:delText>${original}</w:delText></w:r></w:del>`,
    `  <w:ins w:id="${insertId}" w:author="A" w:date="2026-06-01T09:30:00Z"><w:r><w:t>${current}</w:t></w:r></w:ins>`,
    "</w:p>",
  ].join("");
}

function namespacedRevisionParagraph(deleteId: string, insertId: string, original: string, current: string): string {
  return [
    "<w:p>",
    `  <w:del xmlns:w14="${W14_NS}" w:id="${deleteId}" w:author="A" w:date="2026-06-01T09:30:00Z"><w:r><w:rPr><w14:glow w14:rad="50000"/></w:rPr><w:delText>${original}</w:delText></w:r></w:del>`,
    `  <w:ins xmlns:w14="${W14_NS}" w:id="${insertId}" w:author="A" w:date="2026-06-01T09:30:00Z"><w:r><w:rPr><w14:glow w14:rad="50000"/></w:rPr><w:t>${current}</w:t></w:r></w:ins>`,
    "</w:p>",
  ].join("");
}

function extractStoryText(xml: string, includeDeletedText: boolean): string {
  const parsed = parseXml(xml);
  return parsed.elements
    .filter((element) => element.namespaceURI === W_NS && (element.localName === "t" || (includeDeletedText && element.localName === "delText")))
    .map((element) => element.text)
    .join("");
}

function asView(view: string): RevisionView {
  if (view === "current" || view === "original" || view === "all") {
    return view;
  }
  throw new Error(`unknown DOCX revision view ${view}`);
}

function state(context: Record<string, unknown>): RevisionAcceptanceState {
  const current = context.state as RevisionAcceptanceState | undefined;
  if (!current) {
    throw new Error("Missing DOCX revision acceptance state");
  }
  return current;
}

function pkg(context: Record<string, unknown>): OpcPackage {
  return required(state(context).pkg, "missing DOCX revision package");
}

function readStoryTextRows(context: Record<string, unknown>): Array<{ part: string; text: string }> {
  const rows = requiredDataTable(context);
  expect(rows[0]).toEqual(["part", "text"]);
  return rows.slice(1).map((row) => ({
    part: row[0] ?? "",
    text: row[1] ?? "",
  }));
}

function readPartRows(context: Record<string, unknown>): string[] {
  const rows = requiredDataTable(context);
  expect(rows[0]).toEqual(["part"]);
  return rows.slice(1).map((row) => row[0] ?? "");
}

function requiredDataTable(context: Record<string, unknown>): string[][] {
  return required(
    (context.step as { argument?: { dataTable?: string[][] } } | undefined)?.argument?.dataTable,
    "missing data table argument",
  );
}

function required<T>(value: T | undefined, message: string): T {
  expect(value).toBeDefined();
  if (value === undefined) {
    throw new Error(message);
  }
  return value;
}
