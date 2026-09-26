import { expect } from "bun:test";
import { readFileSync } from "node:fs";
import { join, resolve } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { addPart, addRelationship, type OpcPackage } from "../../src/opc/index.ts";
import { inspectStories, storyParts, type RevisionView, type StoryInspection } from "../../src/docx/story.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const PYTHON_WORD_ROOT = join(
  PROJECT_ROOT,
  "fixtures/python-office-mcp-server/tests/_templates/testdata/word",
);
const UPSTREAM_FIXTURES = join(
  PROJECT_ROOT,
  "references/fixtures-ooxml/reference-assets/docx/tests/paper/fixtures/generated",
);
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const FOOTNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const ENDNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

type StoryAcceptanceState = {
  pkg?: OpcPackage;
  inspection?: StoryInspection;
  listedParts?: ReturnType<typeof storyParts>;
  rememberedBytes?: Uint8Array;
};

export const bindings: StepBinding[] = [
  {
    pattern: /^DOCX story fixture "([^"]+)" is opened$/,
    run: async (context, name) => {
      state(context).inspection = undefined;
      state(context).listedParts = undefined;
      state(context).rememberedBytes = undefined;
      state(context).pkg = await openStoryFixture(name);
    },
  },
  {
    pattern: /^DOCX story package bytes are remembered$/,
    run: (context) => {
      state(context).rememberedBytes = pkg(context).toBytes();
    },
  },
  {
    pattern: /^DOCX story parts are listed$/,
    run: (context) => {
      state(context).listedParts = storyParts(pkg(context));
    },
  },
  {
    pattern: /^DOCX stories are inspected in "([^"]+)" view$/,
    run: (context, view) => {
      state(context).inspection = inspectStories(pkg(context), { view: view as RevisionView });
    },
  },
  {
    pattern: /^DOCX story parts equal:$/,
    run: (context) => {
      const actual = required(state(context).listedParts, "missing DOCX story parts");
      expect(actual as Array<{part:string;kind:string}>).toEqual(readPartRows(context));
    },
  },
  {
    pattern: /^DOCX story "([^"]+)" paragraphs equal:$/,
    run: (context, part) => {
      const actual = required(
        required(state(context).inspection, "missing DOCX story inspection").stories.find((story) => story.part === part),
        `missing inspected story ${part}`,
      );
      expect(actual.paragraphs).toEqual(readParagraphRows(context));
    },
  },
  {
    pattern: /^DOCX blind regions equal:$/,
    run: (context) => {
      const actual = required(state(context).inspection, "missing DOCX story inspection").blindRegions;
      expect(actual).toEqual(readBlindRows(context));
    },
  },
  {
    pattern: /^DOCX story package bytes equal the remembered bytes$/,
    run: (context) => {
      expect(pkg(context).toBytes()).toEqual(required(state(context).rememberedBytes, "missing remembered DOCX bytes"));
    },
  },
];

export async function openStoryFixture(name: string): Promise<OpcPackage> {
  if (name === "synthetic-revisions") {
    return createSyntheticRevisionFixture();
  }
  if (name === "synthetic-blind") {
    return createSyntheticBlindFixture();
  }
  return (await import("../../src/opc/index.ts")).OpcPackage.open(resolveStoryFixturePath(name));
}

export function resolveStoryFixturePath(name: string): string {
  const real = new Map<string, string>([
    ["gauntlet", join(UPSTREAM_FIXTURES, "gauntlet/gauntlet.docx")],
    ["header-footer-sections", join(UPSTREAM_FIXTURES, "feature-isolated/header-footer-sections.docx")],
    ["footnotes-endnotes", join(UPSTREAM_FIXTURES, "feature-isolated/footnotes-endnotes.docx")],
    ["comments", join(UPSTREAM_FIXTURES, "feature-isolated/comments.docx")],
    ["textbox", join(UPSTREAM_FIXTURES, "feature-isolated/textbox.docx")],
  ]);
  return required(real.get(name), `unknown DOCX story fixture ${name}`);
}

export async function createSyntheticRevisionFixture(): Promise<OpcPackage> {
  const { OpcPackage } = await import("../../src/opc/index.ts");
  const pkg = await OpcPackage.open(new Uint8Array(readFileSync(join(PYTHON_WORD_ROOT, "single_paragraph.docx"))));
  const main = pkg.mainPart();

  pkg.set(main, [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}">`,
    "  <w:body>",
    "    <w:p><w:r><w:t>Body intro</w:t></w:r></w:p>",
    "    <w:tbl>",
    "      <w:tr><w:tc><w:p>",
    "        <w:r><w:t>Table</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Cell</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Line2</w:t></w:r>",
    "      </w:p></w:tc></w:tr>",
    "    </w:tbl>",
    "    <w:ins><w:p><w:r><w:t>Inserted body paragraph</w:t></w:r></w:p></w:ins>",
    "    <w:del><w:p><w:r><w:delText>Deleted body paragraph</w:delText></w:r></w:p></w:del>",
    "    <w:p><w:r><w:t>Body kept</w:t></w:r></w:p>",
    "    <w:sectPr/>",
    "  </w:body>",
    "</w:document>",
  ].join("\n"));

  addPart(pkg, "word/header1.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<wx:hdr xmlns:wx="${W_NS}">`,
    "  <wx:p>",
    "    <wx:r><wx:t>Header </wx:t></wx:r>",
    "    <wx:ins><wx:r><wx:t>new</wx:t></wx:r></wx:ins>",
    "    <wx:del><wx:r><wx:delText>old</wx:delText></wx:r></wx:del>",
    "    <wx:r><wx:t> text</wx:t></wx:r>",
    "  </wx:p>",
    "</wx:hdr>",
  ].join("\n"), HEADER_CONTENT_TYPE);
  addPart(pkg, "word/footer1.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<wf:ftr xmlns:wf="${W_NS}">`,
    "  <wf:p>",
    "    <wf:r><wf:t>Footer </wf:t></wf:r>",
    "    <wf:ins><wf:r><wf:t>current</wf:t></wf:r></wf:ins>",
    "    <wf:del><wf:r><wf:delText>original</wf:delText></wf:r></wf:del>",
    "  </wf:p>",
    "</wf:ftr>",
  ].join("\n"), FOOTER_CONTENT_TYPE);
  addPart(pkg, "word/footer2.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<alt:ftr xmlns:alt="${W_NS}"><alt:p><alt:r><alt:t>Footer via header relationship</alt:t></alt:r></alt:p></alt:ftr>`,
  ].join("\n"), FOOTER_CONTENT_TYPE);
  addPart(pkg, "word/footnotes.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<fn:footnotes xmlns:fn="${W_NS}">`,
    "  <fn:footnote fn:id=\"-1\" fn:type=\"separator\"><fn:p><fn:r><fn:separator/></fn:r></fn:p></fn:footnote>",
    "  <fn:footnote fn:id=\"1\">",
    "    <fn:p>",
    "      <fn:r><fn:t xml:space=\"preserve\"> </fn:t></fn:r>",
    "      <fn:ins><fn:r><fn:t>Footnote new</fn:t></fn:r></fn:ins>",
    "      <fn:del><fn:r><fn:delText>Footnote old</fn:delText></fn:r></fn:del>",
    "      <fn:r><fn:t xml:space=\"preserve\"> text</fn:t></fn:r>",
    "    </fn:p>",
    "  </fn:footnote>",
    "</fn:footnotes>",
  ].join("\n"), FOOTNOTES_CONTENT_TYPE);
  addPart(pkg, "word/endnotes.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<en:endnotes xmlns:en="${W_NS}">`,
    "  <en:endnote en:id=\"-1\" en:type=\"separator\"><en:p><en:r><en:separator/></en:r></en:p></en:endnote>",
    "  <en:endnote en:id=\"1\"><en:p><en:r><en:t xml:space=\"preserve\"> Endnote text</en:t></en:r></en:p></en:endnote>",
    "</en:endnotes>",
  ].join("\n"), ENDNOTES_CONTENT_TYPE);
  addPart(pkg, "word/comments.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<cm:comments xmlns:cm="${W_NS}">`,
    "  <cm:comment cm:id=\"0\" cm:author=\"A\">",
    "    <cm:p>",
    "      <cm:r><cm:t>Comment </cm:t></cm:r>",
    "      <cm:ins><cm:r><cm:t>current</cm:t></cm:r></cm:ins>",
    "      <cm:del><cm:r><cm:delText>original</cm:delText></cm:r></cm:del>",
    "    </cm:p>",
    "  </cm:comment>",
    "</cm:comments>",
  ].join("\n"), COMMENTS_CONTENT_TYPE);

  addRelationship(pkg, main, `${OFFICE_REL}header`, "header1.xml");
  addRelationship(pkg, main, `${OFFICE_REL}header`, "header1.xml", { id: "rId42" });
  addRelationship(pkg, main, `${OFFICE_REL}footer`, "footer1.xml");
  addRelationship(pkg, main, `${OFFICE_REL}footnotes`, "footnotes.xml");
  addRelationship(pkg, main, `${OFFICE_REL}endnotes`, "endnotes.xml");
  addRelationship(pkg, main, `${OFFICE_REL}comments`, "comments.xml");
  addRelationship(pkg, "word/header1.xml", `${OFFICE_REL}footer`, "footer2.xml");

  return pkg;
}

export async function createSyntheticBlindFixture(): Promise<OpcPackage> {
  const { OpcPackage } = await import("../../src/opc/index.ts");
  const pkg = await OpcPackage.open(new Uint8Array(readFileSync(join(PYTHON_WORD_ROOT, "single_paragraph.docx"))));
  const main = pkg.mainPart();

  pkg.set(main, [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:document xmlns:w="${W_NS}" xmlns:ux="urn:test:opaque" xmlns:alt="${MC_NS}">`,
    "  <w:body>",
    "    <w:p>",
    "      <w:r><w:t>Field </w:t></w:r>",
    "      <w:fldSimple w:instr=\"DATE\"><w:r><w:t>2026-01-01</w:t></w:r></w:fldSimple>",
    "    </w:p>",
    "    <w:p>",
    "      <w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>",
    "      <w:r><w:instrText> PAGE </w:instrText></w:r>",
    "      <w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>",
    "      <w:r><w:t>7</w:t></w:r>",
    "      <w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
    "    </w:p>",
    "    <w:p>",
    "      <ux:opaque><w:r><w:t>Opaque text</w:t></w:r></ux:opaque>",
    "    </w:p>",
    "    <w:p>",
    "      <alt:AlternateContent>",
    `        <alt:Choice Requires=\"wps\" xmlns:wps=\"urn:test:wps\"><w:r><w:t>Choice text</w:t></w:r></alt:Choice>`,
    "        <alt:Fallback><w:r><w:t>Fallback text</w:t></w:r></alt:Fallback>",
    "      </alt:AlternateContent>",
    "    </w:p>",
    "    <w:p>",
    "      <w:r><w:t>Outside</w:t></w:r>",
    "      <w:txbxContent><w:p><w:r><w:t>Text box text</w:t></w:r></w:p></w:txbxContent>",
    "    </w:p>",
    "    <w:sectPr/>",
    "  </w:body>",
    "</w:document>",
  ].join("\n"));

  addPart(pkg, "word/header9.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<orphan:hdr xmlns:orphan="${W_NS}"><orphan:p><orphan:r><orphan:t>Orphaned header</orphan:t></orphan:r></orphan:p></orphan:hdr>`,
  ].join("\n"), HEADER_CONTENT_TYPE);
  addRelationship(pkg, main, `${OFFICE_REL}header`, "https://example.test/external-header.xml", { external: true });

  return pkg;
}

function state(context: Record<string, unknown>): StoryAcceptanceState {
  const state = context.state as StoryAcceptanceState | undefined;
  if (!state) {
    throw new Error("Missing DOCX story acceptance state");
  }
  return state;
}

function pkg(context: Record<string, unknown>): OpcPackage {
  return required(state(context).pkg, "missing DOCX story package");
}

function required<T>(value: T | undefined, message: string): T {
  expect(value).toBeDefined();
  if (value === undefined) {
    throw new Error(message);
  }
  return value;
}

function readPartRows(context: Record<string, unknown>): Array<{ part: string; kind: string }> {
  const rows = requiredDataTable(context);
  expect(rows[0]).toEqual(["part", "kind"]);
  return rows.slice(1).map((row) => ({
    part: row[0] ?? "",
    kind: row[1] ?? "",
  }));
}

function readParagraphRows(context: Record<string, unknown>): Array<{ index: number; text: string }> {
  const rows = requiredDataTable(context);
  expect(rows[0]).toEqual(["index", "text"]);
  return rows.slice(1).map((row) => ({
    index: Number.parseInt(row[0] ?? "", 10),
    text: decodeEscapes(row[1] ?? ""),
  }));
}

function readBlindRows(context: Record<string, unknown>): Array<{ part: string; kind: string; count: number }> {
  const rows = requiredDataTable(context);
  expect(rows[0]).toEqual(["part", "kind", "count"]);
  return rows.slice(1).map((row) => ({
    part: row[0] ?? "",
    kind: row[1] ?? "",
    count: Number.parseInt(row[2] ?? "", 10),
  }));
}

function requiredDataTable(context: Record<string, unknown>): string[][] {
  const table = (context.step as { argument?: { dataTable?: string[][] } } | undefined)?.argument?.dataTable;
  return required(table, "missing data table argument");
}

function decodeEscapes(text: string): string {
  return text
    .replaceAll("\\\\", "\u0000")
    .replaceAll("\\u0020", " ")
    .replaceAll("\\t", "\t")
    .replaceAll("\\n", "\n")
    .replaceAll("\u0000", "\\");
}
