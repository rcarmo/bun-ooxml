import {fixturePaths,GO_TESTDATA_PACKAGE_IDS,PYTHON_TESTDATA_PACKAGE_IDS} from "../../scripts/fixture-inputs.ts";
import assert from "node:assert/strict";
import { readdirSync } from "node:fs";
import { join, resolve } from "node:path";
import { deflateRawSync } from "node:zlib";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { OoxmlError } from "../../src/errors.ts";
import { OpcPackage } from "../../src/opc/package.ts";
import { crc32, readZip, writeZip } from "../../src/opc/zip.ts";
import { applyEdits, escapeText, parseXml, type XmlDocument, type XmlElement } from "../../src/xml/index.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8", { fatal: true });

const LOCAL_SIGNATURE = 0x04034b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;

const FLAG_ENCRYPTED = 0x0001;
const FLAG_DATA_DESCRIPTOR = 0x0008;
const FLAG_UTF8 = 0x0800;

const METHOD_STORED = 0;
const METHOD_DEFLATED = 8;

const MAIN_XML_PART = "doc/main.xml";
const PAYLOAD_PART = "custom/data.bin";

const RICH_XML = [
  '<?xml version="1.0"?>',
  '<!--before-->',
  '<?pi ok?>',
  '<ns:root xmlns="urn:default" xmlns:ns="urn:ns" xmlns:x="urn:x" a="1 &amp; 2">',
  'pre',
  '<x:child x:b="v"/>',
  'mid',
  '<![CDATA[<tail>]]>',
  '</ns:root>',
  '<!--after-->',
].join("");

const MALFORMED_XML_CASES: XmlCase[] = [
  {
    label: "malformed declaration",
    xml: '<?xml version="1.0" extra="x"?><r/>',
    expectedCode: "XML_MALFORMED",
  },
  {
    label: "malformed processing instruction",
    xml: "<?pi/?>",
    expectedCode: "XML_MALFORMED",
  },
  {
    label: "invalid comment termination",
    xml: "<r><!--bad---></r>",
    expectedCode: "XML_MALFORMED",
  },
  {
    label: "missing attribute whitespace",
    xml: "<r a='1'b='2'/>",
    expectedCode: "XML_MALFORMED",
  },
  {
    label: "reserved namespace misuse",
    xml: '<r xmlns="http://www.w3.org/XML/1998/namespace"/>',
    expectedCode: "XML_MALFORMED",
  },
  {
    label: "DTD",
    xml: "<!DOCTYPE r><r/>",
    expectedCode: "XML_DTD_FORBIDDEN",
  },
  {
    label: "undeclared entity",
    xml: "<r>&custom;</r>",
    expectedCode: "XML_ENTITY_FORBIDDEN",
  },
  {
    label: "duplicate attribute",
    xml: "<r a='1' a='2'/>",
    expectedCode: "XML_DUPLICATE_ATTRIBUTE",
  },
  {
    label: "unbound prefix",
    xml: "<x:r/>",
    expectedCode: "XML_UNBOUND_PREFIX",
  },
  {
    label: "mismatched tag",
    xml: "<a></b>",
    expectedCode: "XML_MISMATCHED_TAG",
  },
];

const BOUNDS_XML_CASES: XmlCase[] = [
  {
    label: "depth",
    xml: `${"<n>".repeat(257)}${"</n>".repeat(257)}`,
    expectedCode: "XML_DEPTH_LIMIT",
  },
  {
    label: "nodes",
    xml: `<r>${"<n/>".repeat(100_000)}</r>`,
    expectedCode: "XML_NODE_LIMIT",
  },
  {
    label: "length",
    xml: `<r>${"a".repeat(8 * 1024 * 1024)}</r>`,
    expectedCode: "XML_INPUT_TOO_LARGE",
  },
];

type XmlCase = {
  label: string;
  xml: string;
  expectedCode: string;
};

type XmlParseResult = {
  label: string;
  expectedCode?: string;
  document?: XmlDocument;
  error?: unknown;
};

type XmlState = {
  mode?: "success" | "refusal" | "bounds" | "edits";
  cases?: XmlCase[];
  results?: XmlParseResult[];
  editSource?: string;
  editedXml?: string;
  editRefusals?: unknown[];
};

type ZipUnsafeCase = {
  name: string;
  archive: Uint8Array;
  expectedCode: string;
};

type ZipState = {
  mode?: "valid" | "unsafe" | "bounds" | "write";
  validMembers?: ZipMemberSpec[];
  archive?: Uint8Array;
  expectedOrder?: string[];
  expectedPayloads?: Map<string, Uint8Array>;
  readParts?: Map<string, Uint8Array>;
  unsafeCases?: ZipUnsafeCase[];
  unsafeResults?: Map<string, unknown>;
  boundsResults?: Array<{ name: string; expectedCode: string; error: unknown }>;
  writeInput?: Map<string, Uint8Array>;
  writeFirst?: Uint8Array;
  writeSecond?: Uint8Array;
  writeDescription?: ZipDescription;
  collisionError?: unknown;
};

type CorpusInfo = {
  name: string;
  root: string;
  allFiles: string[];
  packageFiles: string[];
};

type CorpusRoundTrip = {
  path: string;
  noOpBytesMatch: boolean;
  reopenedBytesMatch: boolean;
};

type OpcState = {
  corpora?: CorpusInfo[];
  corpusRoundTrips?: CorpusRoundTrip[];
  rollbackError?: unknown;
  rollbackOriginalBytes?: Uint8Array;
  rollbackAfterBytes?: Uint8Array;
  rollbackOriginalMain?: Uint8Array;
  rollbackCurrentMain?: Uint8Array;
  rollbackOriginalPayload?: Uint8Array;
  rollbackCurrentPayload?: Uint8Array;
  updatedXmlText?: string;
  originalPayload?: Uint8Array;
  reopenedPayload?: Uint8Array;
};

type CoreState = {
  xml?: XmlState;
  zip?: ZipState;
  opc?: OpcState;
};

type RootState = {
  core?: CoreState;
};

type ZipMemberSpec = {
  name: string;
  blob?: Uint8Array;
  method?: number;
  flags?: number;
  data?: Uint8Array;
  crc?: number;
  fileSize?: number;
  compressedSize?: number;
  localName?: string;
  localFlags?: number;
  localMethod?: number;
  localCrc?: number;
  localFileSize?: number;
  localCompressedSize?: number;
  localExtra?: Uint8Array;
  centralExtra?: Uint8Array;
  comment?: Uint8Array;
  headerOffset?: number;
  diskStart?: number;
  dataDescriptor?: boolean;
  descriptorSignature?: boolean;
  descriptorCrc?: number;
  descriptorFileSize?: number;
  descriptorCompressedSize?: number;
};

type ZipBuildOptions = {
  diskNumber?: number;
  centralDisk?: number;
  diskEntries?: number;
  totalEntries?: number;
  centralSizeDelta?: number;
  archiveComment?: Uint8Array;
};

type ZipCentralRecord = {
  name: string;
  flags: number;
  method: number;
  headerOffset: number;
};

type ZipLocalRecord = {
  name: string;
  flags: number;
  method: number;
};

type ZipDescription = {
  central: ZipCentralRecord[];
  local: ZipLocalRecord[];
  totalEntries: number;
  centralSize: number;
  centralOffset: number;
};

export const bindings: StepBinding[] = [
  {
    pattern: /^XML text and attributes containing raw CRLF and character references$/,
    run: context => { context.normalisedSource = '<r a="x\r\ny\tz&#xD;&#xA;&#x9;">u\r\nv\rw&#xD;<![CDATA[c\r\nd]]><s/></r>'; },
  },
  {
    pattern: /^that XML is parsed without rewriting the source$/,
    run: context => { context.normalisedDoc = parseXml(context.normalisedSource as string); },
  },
  {
    pattern: /^decoded text normalises raw line endings but preserves referenced carriage returns$/,
    run: context => { assert.equal((context.normalisedDoc as ReturnType<typeof parseXml>).root.text, 'u\nv\nw\rc\nd'); },
  },
  {
    pattern: /^decoded attributes normalise literal whitespace while preserving referenced whitespace$/,
    run: context => { assert.equal((context.normalisedDoc as ReturnType<typeof parseXml>).root.attributes.a, 'x y z\r\n\t'); },
  },
  {
    pattern: /^element offsets still address the original source string$/,
    run: context => { const child = (context.normalisedDoc as ReturnType<typeof parseXml>).root.children[0]!; assert.equal((context.normalisedSource as string).slice(child.start,child.end), '<s/>'); },
  },
  {
    pattern: /^an XML document with a declaration, comments, processing instructions and namespaces$/,
    run: (context) => {
      const state = coreState(context);
      state.xml = {
        mode: "success",
        cases: [{ label: "rich", xml: RICH_XML, expectedCode: "" }],
      };
    },
  },
  {
    pattern: /^XML containing a malformed declaration or processing instruction$/,
    run: (context) => {
      const state = coreState(context);
      state.xml = {
        mode: "refusal",
        cases: MALFORMED_XML_CASES.slice(0, 2),
      };
    },
  },
  {
    pattern: /^XML containing invalid comment termination or missing attribute whitespace$/,
    run: (context) => {
      const xml = requireXmlState(coreState(context), "refusal");
      xml.cases?.push(...MALFORMED_XML_CASES.slice(2, 4));
    },
  },
  {
    pattern: /^XML containing reserved namespace misuse, a DTD or an undeclared entity$/,
    run: (context) => {
      const xml = requireXmlState(coreState(context), "refusal");
      xml.cases?.push(...MALFORMED_XML_CASES.slice(4, 7));
    },
  },
  {
    pattern: /^XML containing a duplicate attribute, an unbound prefix or a mismatched tag$/,
    run: (context) => {
      const xml = requireXmlState(coreState(context), "refusal");
      xml.cases?.push(...MALFORMED_XML_CASES.slice(7));
    },
  },
  {
    pattern: /^the document is parsed$/,
    run: (context) => {
      const xml = requireDefined(coreState(context).xml, "Missing XML scenario state");
      const cases = requireDefined(xml.cases, "Missing XML cases");
      xml.results = cases.map((item) => {
        try {
          return {
            label: item.label,
            expectedCode: item.expectedCode || undefined,
            document: parseXml(item.xml),
          } satisfies XmlParseResult;
        } catch (error) {
          return {
            label: item.label,
            expectedCode: item.expectedCode || undefined,
            error,
          } satisfies XmlParseResult;
        }
      });
    },
  },
  {
    pattern: /^the root element and descendants expose decoded text, decoded attributes and namespace URIs$/,
    run: (context) => {
      const result = singleParsedDocument(coreState(context));
      const root = result.root;
      const child = requireDefined(root.children[0], "Missing XML child element");
      assert.equal(result.elements.length, 2);
      assert.equal(root.name, "ns:root");
      assert.equal(root.localName, "root");
      assert.equal(root.namespaceURI, "urn:ns");
      assert.deepEqual({ ...root.attributes }, {
        xmlns: "urn:default",
        "xmlns:ns": "urn:ns",
        "xmlns:x": "urn:x",
        a: "1 & 2",
      });
      assert.equal(root.text, "premid<tail>");
      assert.equal(child.name, "x:child");
      assert.equal(child.localName, "child");
      assert.equal(child.namespaceURI, "urn:x");
      assert.deepEqual({ ...child.attributes }, { "x:b": "v" });
      assert.equal(child.text, "");
    },
  },
  {
    pattern: /^each element exposes UTF-16 source offsets, parent links, child links, root links and self-closing state$/,
    run: (context) => {
      const result = singleParsedDocument(coreState(context));
      const root = result.root;
      const child = requireDefined(root.children[0], "Missing XML child element");
      assert.equal(root.start, RICH_XML.indexOf("<ns:root"));
      assert.equal(root.openEnd, RICH_XML.indexOf(">", root.start) + 1);
      assert.equal(root.closeStart, RICH_XML.lastIndexOf("</ns:root>"));
      assert.equal(root.end, root.closeStart + "</ns:root>".length);
      assert.equal(root.parent, undefined);
      assert.equal(root.root, root);
      assert.equal(root.selfClosing, false);
      assert.equal(root.children.length, 1);
      assert.equal(child.start, RICH_XML.indexOf("<x:child"));
      assert.equal(child.openEnd, RICH_XML.indexOf("/>", child.start) + 2);
      assert.equal(child.closeStart, child.openEnd);
      assert.equal(child.end, child.openEnd);
      assert.equal(child.parent, root);
      assert.equal(child.root, root);
      assert.equal(child.selfClosing, true);
      assert.equal(child.children.length, 0);
    },
  },
  {
    pattern: /^XML whose nesting depth, node count or input length exceeds the configured parser limits$/,
    run: (context) => {
      const state = coreState(context);
      state.xml = {
        mode: "bounds",
        cases: [...BOUNDS_XML_CASES],
      };
    },
  },
  {
    pattern: /^(?:parsing is refused with a stable XML error code|parsing is refused before returning a partial tree)$/, 
    run: (context) => {
      const xml = requireDefined(coreState(context).xml, "Missing XML scenario state");
      const results = requireDefined(xml.results, "Missing XML parse results");
      assert.ok(results.length > 0, "Expected XML parse results");
      for (const result of results) {
        assert.equal(result.document, undefined, `${result.label} unexpectedly parsed`);
        assert.equal(refusalCode(result.error), result.expectedCode, `${result.label} refusal code mismatch`);
      }
    },
  },
  {
    pattern: /^a well-formed XML document and source offsets for text or element content$/,
    run: (context) => {
      const state = coreState(context);
      state.xml = {
        mode: "edits",
        editSource: "<r><slot>two</slot>one</r>",
        editRefusals: [],
      };
    },
  },
  {
    pattern: /^disjoint edits are applied with escaped replacement text or XML fragments$/,
    run: (context) => {
      const xml = requireXmlState(coreState(context), "edits");
      const source = requireDefined(xml.editSource, "Missing XML edit source");
      const parsed = parseXml(source);
      const slot = requireDefined(parsed.root.children[0], "Missing <slot> element");
      const textStart = source.lastIndexOf("one");
      xml.editedXml = applyEdits(source, [
        {
          start: slot.openEnd,
          end: slot.closeStart,
          value: "<x/>",
        },
        {
          start: textStart,
          end: textStart + 3,
          value: escapeText("1 < 2"),
        },
      ]);
      xml.editRefusals = [
        capture(() =>
          applyEdits(source, [
            { start: slot.openEnd, end: slot.closeStart, value: "<x/>" },
            { start: slot.openEnd + 1, end: slot.closeStart, value: "<y/>" },
          ])),
        capture(() =>
          applyEdits(source, [
            { start: slot.openEnd, end: slot.closeStart, value: "<x>" },
          ])),
        capture(() =>
          applyEdits(source, [
            { start: slot.openEnd, end: slot.closeStart, value: "<!DOCTYPE x><x/>" },
          ])),
      ];
    },
  },
  {
    pattern: /^the resulting XML stays well formed and DTD free$/,
    run: (context) => {
      const xml = requireXmlState(coreState(context), "edits");
      const edited = requireDefined(xml.editedXml, "Missing edited XML");
      const parsed = parseXml(edited);
      const slot = requireDefined(parsed.root.children[0], "Missing slot after edit");
      assert.equal(parsed.root.localName, "r");
      assert.equal(parsed.root.text, "1 < 2");
      assert.equal(slot.localName, "slot");
      assert.equal(slot.children.length, 1);
      assert.equal(slot.text, "");
      assert.equal(requireDefined(slot.children[0], "Missing inserted XML fragment").localName, "x");
      assert.equal(edited.includes("<!DOCTYPE"), false);
    },
  },
  {
    pattern: /^overlapping edits or edits that leave malformed or DTD-bearing XML are refused before returning changed text$/,
    run: (context) => {
      const xml = requireXmlState(coreState(context), "edits");
      const refusals = requireDefined(xml.editRefusals, "Missing XML edit refusals");
      assert.equal(refusalCode(refusals[0]), "XML_EDIT_OVERLAP");
      assert.equal(refusalCode(refusals[1]), "XML_EDIT_UNSAFE");
      assert.equal(refusalCode(refusals[2]), "XML_EDIT_UNSAFE");
    },
  },
  {
    pattern: /^a ZIP archive with canonical OPC member names$/,
    run: (context) => {
      const state = coreState(context);
      state.zip = {
        mode: "valid",
        validMembers: [
          { name: "[Content_Types].xml", blob: encoder.encode("<Types/>") },
          { name: "docProps/core.xml", blob: encoder.encode("<core/>") },
          { name: "word/document.xml", blob: encoder.encode("<w:document/>".repeat(24)) },
        ],
        expectedOrder: ["[Content_Types].xml", "docProps/core.xml", "word/document.xml"],
      };
    },
  },
  {
    pattern: /^the archive contains stored and deflated file entries$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "valid");
      const members = requireDefined(zip.validMembers, "Missing valid ZIP members");
      const contentTypes = requireDefined(members.find((member) => member.name === "[Content_Types].xml"));
      const core = requireDefined(members.find((member) => member.name === "docProps/core.xml"));
      const document = requireDefined(members.find((member) => member.name === "word/document.xml"));
      contentTypes.method = METHOD_STORED;
      core.method = METHOD_STORED;
      document.method = METHOD_DEFLATED;
      document.flags = FLAG_UTF8 | FLAG_DATA_DESCRIPTOR;
      document.dataDescriptor = true;
      document.descriptorSignature = true;
    },
  },
  {
    pattern: /^the archive may contain zero-byte directory entries and a declared archive comment$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "valid");
      const files = requireDefined(zip.validMembers, "Missing valid ZIP members");
      const members = [
        { name: "word/", method: METHOD_STORED, blob: new Uint8Array() },
        ...files,
      ];
      zip.archive = buildZip(members, {
        archiveComment: encoder.encode("kept as declared ZIP comment"),
      });
      zip.expectedPayloads = new Map(
        members
          .filter((member) => !member.name.endsWith("/"))
          .map((member) => [member.name, member.blob ?? encoder.encode(member.name)]),
      );
    },
  },
  {
    pattern: /^the archive is read through the shared ZIP module$/,
    run: (context) => {
      const zip = requireDefined(coreState(context).zip, "Missing ZIP scenario state");
      if (zip.mode === "valid") {
        zip.readParts = readZip(requireDefined(zip.archive, "Missing valid ZIP archive"));
        return;
      }
      if (zip.mode === "unsafe") {
        zip.unsafeResults = new Map(
          requireDefined(zip.unsafeCases, "Missing unsafe ZIP cases").map((item) => [
            item.name,
            capture(() => readZip(item.archive)),
          ]),
        );
        return;
      }
      if (zip.mode === "bounds") {
        const archive = requireDefined(zip.archive, "Missing bounded ZIP archive");
        zip.boundsResults = [
          {
            name: "archive size",
            expectedCode: "zip-archive-too-large",
            error: capture(() => readZip(archive, { maxArchiveBytes: archive.length - 1 })),
          },
          {
            name: "entry count",
            expectedCode: "zip-too-many-entries",
            error: capture(() => readZip(archive, { maxEntries: 1 })),
          },
          {
            name: "entry size",
            expectedCode: "zip-entry-too-large",
            error: capture(() => readZip(archive, { maxEntryBytes: 8 })),
          },
          {
            name: "total expanded size",
            expectedCode: "zip-total-too-large",
            error: capture(() => readZip(archive, { maxTotalBytes: 8 })),
          },
          {
            name: "compression ratio",
            expectedCode: "zip-compression-ratio-exceeded",
            error: capture(() => readZip(archive, { maxCompressionRatio: 2 })),
          },
        ];
        return;
      }
      throw new Error(`Unsupported ZIP mode: ${zip.mode ?? "unset"}`);
    },
  },
  {
    pattern: /^file entries are returned in central-directory order$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "valid");
      const names = [...requireDefined(zip.readParts, "Missing ZIP read result").keys()];
      assert.deepEqual(names, requireDefined(zip.expectedOrder, "Missing expected ZIP order"));
    },
  },
  {
    pattern: /^directory entries do not become package parts$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "valid");
      const parts = requireDefined(zip.readParts, "Missing ZIP read result");
      assert.equal(parts.has("word/"), false);
    },
  },
  {
    pattern: /^each returned payload matches its declared CRC and size$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "valid");
      const parts = requireDefined(zip.readParts, "Missing ZIP read result");
      const expected = requireDefined(zip.expectedPayloads, "Missing expected ZIP payloads");
      for (const [name, payload] of expected) {
        const actual = requireDefined(parts.get(name), `Missing ZIP member ${name}`);
        assertBytesEqual(actual, payload, `${name} payload`);
        assert.equal(actual.length, payload.length, `${name} size mismatch`);
        assert.equal(crc32(actual), crc32(payload), `${name} CRC mismatch`);
      }
    },
  },
  {
    pattern: /^a ZIP archive whose structure has no single safe reading$/,
    run: (context) => {
      const state = coreState(context);
      state.zip = {
        mode: "unsafe",
        unsafeCases: [
          {
            name: "duplicate",
            expectedCode: "zip-duplicate-entry",
            archive: buildZip([
              { name: "word/document.xml", blob: encoder.encode("one") },
              { name: "word/document.xml", blob: encoder.encode("two") },
            ]),
          },
          {
            name: "case-collision",
            expectedCode: "zip-case-collision",
            archive: buildZip([
              { name: "word/document.xml", blob: encoder.encode("one") },
              { name: "WORD/document.xml", blob: encoder.encode("two") },
            ]),
          },
          {
            name: "noncanonical",
            expectedCode: "zip-name-invalid",
            archive: buildZip([{ name: "../word/document.xml", blob: encoder.encode("bad") }]),
          },
          {
            name: "encrypted",
            expectedCode: "zip-encryption-unsupported",
            archive: buildZip([
              { name: "word/document.xml", blob: encoder.encode("secret"), flags: FLAG_UTF8 | FLAG_ENCRYPTED },
            ]),
          },
          {
            name: "unsupported-method",
            expectedCode: "zip-method-unsupported",
            archive: buildZip([{ name: "word/document.xml", blob: encoder.encode("x"), method: 12 }]),
          },
          {
            name: "multi-disk",
            expectedCode: "zip-multi-disk-unsupported",
            archive: buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }], { diskNumber: 1 }),
          },
          {
            name: "zip64",
            expectedCode: "zip-zip64-unsupported",
            archive: buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }], {
              totalEntries: 0xffff,
              diskEntries: 0xffff,
            }),
          },
          {
            name: "local-mismatch",
            expectedCode: "zip-local-metadata-mismatch",
            archive: buildZip([
              { name: "word/document.xml", localName: "word/other.xml", blob: encoder.encode("x") },
            ]),
          },
          {
            name: "crc",
            expectedCode: "zip-crc-mismatch",
            archive: buildZip([
              { name: "word/document.xml", blob: encoder.encode("payload"), crc: 0xdeadbeef },
            ]),
          },
          {
            name: "size",
            expectedCode: "zip-size-mismatch",
            archive: buildZip([
              {
                name: "word/document.xml",
                blob: encoder.encode("payload"),
                method: METHOD_STORED,
                fileSize: 99,
                compressedSize: 99,
                localFileSize: 99,
                localCompressedSize: 99,
              },
            ]),
          },
          {
            name: "trailing-structure",
            expectedCode: "zip-end-record-missing",
            archive: concatBytes(
              buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }]),
              encoder.encode("\n"),
            ),
          },
        ],
      };
    },
  },
  {
    pattern: /^the module refuses duplicate member names$/,
    run: (context) => {
      assertZipRefusal(coreState(context), "duplicate", "zip-duplicate-entry");
    },
  },
  {
    pattern: /^the module refuses ASCII case-colliding member names$/,
    run: (context) => {
      assertZipRefusal(coreState(context), "case-collision", "zip-case-collision");
    },
  },
  {
    pattern: /^the module refuses noncanonical member paths$/,
    run: (context) => {
      assertZipRefusal(coreState(context), "noncanonical", "zip-name-invalid");
    },
  },
  {
    pattern: /^the module refuses encrypted or unsupported-compression members$/,
    run: (context) => {
      assertZipRefusal(coreState(context), "encrypted", "zip-encryption-unsupported");
      assertZipRefusal(coreState(context), "unsupported-method", "zip-method-unsupported");
    },
  },
  {
    pattern: /^the module refuses multi-disk archives and ZIP64 sentinels without valid end records$/,
    run: (context) => {
      assertZipRefusal(coreState(context), "multi-disk", "zip-multi-disk-unsupported");
      assertZipRefusal(coreState(context), "zip64", "zip-structure-invalid");
    },
  },
  {
    pattern: /^the module refuses local-header metadata that disagrees with the central directory$/,
    run: (context) => {
      assertZipRefusal(coreState(context), "local-mismatch", "zip-local-metadata-mismatch");
    },
  },
  {
    pattern: /^the module refuses CRC failures, size mismatches, and undeclared trailing structure$/,
    run: (context) => {
      assertZipRefusal(coreState(context), "crc", "zip-crc-mismatch");
      assertZipRefusal(coreState(context), "size", "zip-size-mismatch");
      assertZipRefusal(coreState(context), "trailing-structure", "zip-end-record-missing");
    },
  },
  {
    pattern: /^a ZIP archive whose declared archive size, entry count, entry size,\s+total expanded size, or compression ratio exceeds the configured limit$/,
    run: (context) => {
      const state = coreState(context);
      state.zip = {
        mode: "bounds",
        archive: buildZip([
          { name: "a.bin", blob: encoder.encode("A".repeat(4096)) },
          { name: "b.bin", blob: encoder.encode("bb") },
        ]),
      };
    },
  },
  {
    pattern: /^the module refuses before allocating unbounded output$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "bounds");
      const results = requireDefined(zip.boundsResults, "Missing ZIP bounds results");
      assert.deepEqual(
        results.map((item) => item.expectedCode),
        [
          "zip-archive-too-large",
          "zip-too-many-entries",
          "zip-entry-too-large",
          "zip-total-too-large",
          "zip-compression-ratio-exceeded",
        ],
      );
      for (const item of results) {
        assert.equal(refusalCode(item.error), item.expectedCode, `${item.name} refusal code mismatch`);
      }
    },
  },
  {
    pattern: /^a map of canonical OPC member names and bytes$/,
    run: (context) => {
      const state = coreState(context);
      state.zip = {
        mode: "write",
        writeInput: new Map<string, Uint8Array>([
          ["[Content_Types].xml", encoder.encode("<Types/>")],
          ["custom/data.bin", Uint8Array.from(Array.from({ length: 256 }, (_, index) => index))],
          ["word/document.xml", encoder.encode("<w:document>" + "A".repeat(2048) + "</w:document>")],
        ]),
      };
    },
  },
  {
    pattern: /^the map is written through the shared ZIP module twice$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "write");
      const input = requireDefined(zip.writeInput, "Missing ZIP writer input");
      zip.writeFirst = writeZip(input);
      zip.writeSecond = writeZip(input);
      zip.writeDescription = describeZip(zip.writeFirst);
      zip.collisionError = capture(() =>
        writeZip(
          new Map<string, Uint8Array>([
            ["word/document.xml", encoder.encode("one")],
            ["WORD/document.xml", encoder.encode("two")],
          ]),
        ));
    },
  },
  {
    pattern: /^both outputs are byte-identical ZIP32 archives$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "write");
      const first = requireDefined(zip.writeFirst, "Missing first ZIP output");
      const second = requireDefined(zip.writeSecond, "Missing second ZIP output");
      const description = requireDefined(zip.writeDescription, "Missing ZIP description");
      assertBytesEqual(first, second, "deterministic ZIP output");
      assert.ok(description.totalEntries < 0xffff, "ZIP32 total entry count expected");
      assert.ok(description.centralSize < 0xffffffff, "ZIP32 central size expected");
      assert.ok(description.centralOffset < 0xffffffff, "ZIP32 central offset expected");
      assert.deepEqual([...readZip(first).keys()], ["[Content_Types].xml", "custom/data.bin", "word/document.xml"]);
    },
  },
  {
    pattern: /^file entries use stored or deflated encoding$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "write");
      const methods = requireDefined(zip.writeDescription, "Missing ZIP description").central.map((entry) => entry.method);
      assert.ok(methods.every((method) => method === METHOD_STORED || method === METHOD_DEFLATED));
      assert.ok(methods.includes(METHOD_STORED), "Expected at least one stored member");
      assert.ok(methods.includes(METHOD_DEFLATED), "Expected at least one deflated member");
    },
  },
  {
    pattern: /^names are emitted with the UTF-8 ZIP flag$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "write");
      const description = requireDefined(zip.writeDescription, "Missing ZIP description");
      assert.ok(description.central.every((entry) => (entry.flags & FLAG_UTF8) === FLAG_UTF8));
      assert.ok(description.local.every((entry) => (entry.flags & FLAG_UTF8) === FLAG_UTF8));
    },
  },
  {
    pattern: /^writer input that would collide by ASCII case is refused$/,
    run: (context) => {
      const zip = requireZipState(coreState(context), "write");
      assert.equal(refusalCode(zip.collisionError), "zip-case-collision");
    },
  },
  {
    pattern: /^the go-ooxml and python-office-mcp-server fixture corpora are enumerated$/,
    run: (context) => {
      const state = coreState(context);
      state.opc = {
        corpora: [
          enumerateCorpus("go-ooxml", GO_TESTDATA_PACKAGE_IDS),
          enumerateCorpus("python-office-mcp-server", PYTHON_TESTDATA_PACKAGE_IDS),
        ],
      };
    },
  },
  {
    pattern: /^each OOXML fixture package is opened and serialized without edits through the OPC layer$/,
    run: async (context) => {
      const opc = requireDefined(coreState(context).opc, "Missing OPC scenario state");
      const corpora = requireDefined(opc.corpora, "Missing fixture corpora");
      const results: CorpusRoundTrip[] = [];
      for (const corpus of corpora) {
        for (const path of corpus.packageFiles) {
          const original = new Uint8Array(await Bun.file(path).arrayBuffer());
          const opened = await OpcPackage.open(path);
          const noOp = opened.toBytes();
          const reopened = await OpcPackage.open(noOp);
          const reopenedBytes = reopened.toBytes();
          results.push({
            path,
            noOpBytesMatch: sameBytes(original, noOp),
            reopenedBytesMatch: sameBytes(original, reopenedBytes),
          });
        }
      }
      opc.corpusRoundTrips = results;
    },
  },
  {
    pattern: /^every reopened package matches its original whole-archive bytes$/,
    run: (context) => {
      const opc = requireDefined(coreState(context).opc, "Missing OPC scenario state");
      const results = requireDefined(opc.corpusRoundTrips, "Missing OPC corpus results");
      assert.ok(results.length > 0, "Expected fixture packages to be opened");
      for (const result of results) {
        assert.equal(result.noOpBytesMatch, true, `${result.path} changed on no-op serialization`);
        assert.equal(result.reopenedBytesMatch, true, `${result.path} changed after reopen`);
      }
    },
  },
  {
    pattern: /^both fixture corpora contribute their exact known nonzero fixture counts$/,
    run: (context) => {
      const opc = requireDefined(coreState(context).opc, "Missing OPC scenario state");
      const corpora = requireDefined(opc.corpora, "Missing fixture corpora");
      const go = requireDefined(corpora.find((corpus) => corpus.name === "go-ooxml"));
      const python = requireDefined(corpora.find((corpus) => corpus.name === "python-office-mcp-server"));
      assert.equal(go.allFiles.length, 37);
      assert.equal(go.packageFiles.length, 37);
      assert.equal(python.allFiles.length, 35);
      assert.equal(python.packageFiles.length, 35);
      assert.ok(go.allFiles.length > 0);
      assert.ok(go.packageFiles.length > 0);
      assert.ok(python.allFiles.length > 0);
      assert.ok(python.packageFiles.length > 0);
      assert.equal(go.allFiles.length + python.allFiles.length, 72);
      assert.equal(go.packageFiles.length + python.packageFiles.length, 72);
    },
  },
  {
    pattern: /^a valid OPC package with XML and opaque payload parts$/,
    run: (context) => {
      const state = coreState(context);
      state.opc = {
        rollbackOriginalBytes: buildSyntheticOpcPackage(),
      };
    },
  },
  {
    pattern: /^a transactional edit changes multiple parts and then fails$/,
    run: async (context) => {
      const opc = requireDefined(coreState(context).opc, "Missing OPC scenario state");
      const originalBytes = requireDefined(opc.rollbackOriginalBytes, "Missing rollback source bytes");
      const pkg = await OpcPackage.open(originalBytes);
      opc.rollbackOriginalMain = requireDefined(pkg.get(MAIN_XML_PART), `Missing ${MAIN_XML_PART}`);
      opc.rollbackOriginalPayload = requireDefined(pkg.get(PAYLOAD_PART), `Missing ${PAYLOAD_PART}`);
      try {
        pkg.transaction(() => {
          pkg.set(MAIN_XML_PART, xml(`<main xmlns="urn:acceptance"><value>rolled back</value></main>`));
          pkg.set(PAYLOAD_PART, Uint8Array.from([9, 8, 7, 6]));
          throw new OoxmlError("acceptance-rollback", "synthetic rollback");
        });
      } catch (error) {
        opc.rollbackError = error;
      }
      opc.rollbackAfterBytes = pkg.toBytes();
      opc.rollbackCurrentMain = requireDefined(pkg.get(MAIN_XML_PART), `Missing ${MAIN_XML_PART} after rollback`);
      opc.rollbackCurrentPayload = requireDefined(pkg.get(PAYLOAD_PART), `Missing ${PAYLOAD_PART} after rollback`);
    },
  },
  {
    pattern: /^the package reverts to the original bytes and parts after the refusal$/,
    run: (context) => {
      const opc = requireDefined(coreState(context).opc, "Missing OPC scenario state");
      assert.equal(refusalCode(opc.rollbackError), "acceptance-rollback");
      assertBytesEqual(
        requireDefined(opc.rollbackAfterBytes, "Missing rollback result bytes"),
        requireDefined(opc.rollbackOriginalBytes, "Missing rollback source bytes"),
        "transaction rollback archive",
      );
      assertBytesEqual(
        requireDefined(opc.rollbackCurrentMain, "Missing rolled-back main XML"),
        requireDefined(opc.rollbackOriginalMain, "Missing original main XML"),
        MAIN_XML_PART,
      );
      assertBytesEqual(
        requireDefined(opc.rollbackCurrentPayload, "Missing rolled-back payload"),
        requireDefined(opc.rollbackOriginalPayload, "Missing original payload"),
        PAYLOAD_PART,
      );
    },
  },
  {
    pattern: /^a valid OPC package with a main XML part and an unrelated binary payload$/,
    run: (context) => {
      const state = coreState(context);
      state.opc = {
        rollbackOriginalBytes: buildSyntheticOpcPackage(),
      };
    },
  },
  {
    pattern: /^the main XML part text is changed and the package is reopened$/,
    run: async (context) => {
      const opc = requireDefined(coreState(context).opc, "Missing OPC scenario state");
      const originalBytes = requireDefined(opc.rollbackOriginalBytes, "Missing editable package bytes");
      const pkg = await OpcPackage.open(originalBytes);
      opc.originalPayload = requireDefined(pkg.get(PAYLOAD_PART), `Missing ${PAYLOAD_PART}`);
      const xmlText = pkg.text(MAIN_XML_PART);
      const parsed = parseXml(xmlText);
      const valueElement = requireChildByLocalName(parsed.root, "value");
      const updated = applyEdits(xmlText, [
        {
          start: valueElement.openEnd,
          end: valueElement.closeStart,
          value: escapeText("Updated <value>"),
        },
      ]);
      pkg.set(MAIN_XML_PART, updated);
      const reopened = await OpcPackage.open(pkg.toBytes());
      opc.updatedXmlText = parseXml(reopened.text(MAIN_XML_PART)).root.text;
      opc.reopenedPayload = requireDefined(reopened.get(PAYLOAD_PART), `Missing ${PAYLOAD_PART} after reopen`);
    },
  },
  {
    pattern: /^the edited part contains the new text after reopen$/,
    run: (context) => {
      const opc = requireDefined(coreState(context).opc, "Missing OPC scenario state");
      assert.equal(opc.updatedXmlText, "Updated <value>");
    },
  },
  {
    pattern: /^the unrelated payload bytes remain unchanged$/,
    run: (context) => {
      const opc = requireDefined(coreState(context).opc, "Missing OPC scenario state");
      assertBytesEqual(
        requireDefined(opc.reopenedPayload, "Missing reopened payload"),
        requireDefined(opc.originalPayload, "Missing original payload"),
        PAYLOAD_PART,
      );
    },
  },
];

export default bindings;

function coreState(context: Record<string, unknown>): CoreState {
  const root = context.state as RootState | undefined;
  assert.ok(root, "Missing acceptance root state");
  root.core ??= {};
  return root.core;
}

function requireXmlState(state: CoreState, mode: XmlState["mode"]): XmlState {
  const xml = requireDefined(state.xml, "Missing XML scenario state");
  assert.equal(xml.mode, mode, `Expected XML mode ${String(mode)}`);
  return xml;
}

function singleParsedDocument(state: CoreState): XmlDocument {
  const xml = requireDefined(state.xml, "Missing XML scenario state");
  const result = requireDefined(xml.results?.[0], "Missing parsed XML document result");
  assert.equal(result.error, undefined);
  return requireDefined(result.document, "Missing parsed XML document");
}

function requireZipState(state: CoreState, mode: ZipState["mode"]): ZipState {
  const zip = requireDefined(state.zip, "Missing ZIP scenario state");
  assert.equal(zip.mode, mode, `Expected ZIP mode ${String(mode)}`);
  return zip;
}

function assertZipRefusal(state: CoreState, name: string, expectedCode: string): void {
  const zip = requireZipState(state, "unsafe");
  const results = requireDefined(zip.unsafeResults, "Missing unsafe ZIP results");
  assert.equal(refusalCode(results.get(name)), expectedCode, `${name} ZIP refusal code mismatch`);
}

function refusalCode(error: unknown): string {
  assert.ok(error instanceof OoxmlError, `Expected OoxmlError, got ${String(error)}`);
  return error.code;
}

function capture(action: () => unknown): unknown {
  try {
    action();
  } catch (error) {
    return error;
  }
  throw new Error("Expected action to throw");
}

function requireDefined<T>(value: T | undefined, message = "Missing required value"): T {
  if (value === undefined) throw new Error(message);
  return value;
}

function requireChildByLocalName(element: XmlElement, localName: string): XmlElement {
  return requireDefined(
    element.children.find((child) => child.localName === localName),
    `Missing child element ${localName}`,
  );
}

function enumerateCorpus(name: CorpusInfo["name"], ids: readonly string[]): CorpusInfo { const paths=fixturePaths(ids); return { name, root: "canonical fixture IDs", allFiles: paths, packageFiles: paths }; }

function buildSyntheticOpcPackage(): Uint8Array {
  return writeZip(
    new Map<string, Uint8Array>([
      ["[Content_Types].xml", xml(`
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="bin" ContentType="application/octet-stream"/>
        </Types>
      `)],
      ["_rels/.rels", xml(`
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="doc/main.xml"/>
        </Relationships>
      `)],
      [MAIN_XML_PART, xml(`
        <main xmlns="urn:acceptance">
          <value>Original</value>
        </main>
      `)],
      [PAYLOAD_PART, Uint8Array.from([0, 255, 1, 254, 2, 253])],
    ]),
  );
}

function xml(source: string): Uint8Array {
  return encoder.encode(
    `<?xml version="1.0" encoding="UTF-8"?>${source.trim().replace(/>\s+</g, "><")}`,
  );
}

function buildZip(members: ZipMemberSpec[], options: ZipBuildOptions = {}): Uint8Array {
  const bodyChunks: Uint8Array[] = [];
  const centralChunks: Uint8Array[] = [];
  let bodyLength = 0;

  for (const spec of members) {
    const blob = spec.blob ?? encoder.encode(spec.name);
    const method = spec.method ?? METHOD_DEFLATED;
    const flags = spec.flags ?? FLAG_UTF8;
    const compressed = spec.data ?? compress(method, blob);
    const crc = spec.crc ?? crc32(blob);
    const fileSize = spec.fileSize ?? blob.length;
    const compressedSize = spec.compressedSize ?? compressed.length;
    const localFlags = spec.localFlags ?? flags;
    const localMethod = spec.localMethod ?? method;
    const localCrc = spec.localCrc ?? ((flags & FLAG_DATA_DESCRIPTOR) !== 0 ? 0 : crc);
    const localFileSize = spec.localFileSize ?? ((flags & FLAG_DATA_DESCRIPTOR) !== 0 ? 0 : fileSize);
    const localCompressedSize =
      spec.localCompressedSize ?? ((flags & FLAG_DATA_DESCRIPTOR) !== 0 ? 0 : compressedSize);
    const rawCentralName = encoder.encode(spec.name);
    const rawLocalName = encoder.encode(spec.localName ?? spec.name);
    const localExtra = spec.localExtra ?? new Uint8Array();
    const centralExtra = spec.centralExtra ?? new Uint8Array();
    const comment = spec.comment ?? new Uint8Array();
    const headerOffset = spec.headerOffset ?? bodyLength;
    const descriptor = spec.dataDescriptor
      ? buildDataDescriptor({
          crc: spec.descriptorCrc ?? crc,
          compressedSize: spec.descriptorCompressedSize ?? compressedSize,
          fileSize: spec.descriptorFileSize ?? fileSize,
          withSignature: spec.descriptorSignature ?? false,
        })
      : new Uint8Array();

    bodyChunks.push(
      buildLocalHeader({
        flags: localFlags,
        method: localMethod,
        crc: localCrc,
        compressedSize: localCompressedSize,
        fileSize: localFileSize,
        nameLength: rawLocalName.length,
        extraLength: localExtra.length,
      }),
      rawLocalName,
      localExtra,
      compressed,
      descriptor,
    );
    bodyLength += 30 + rawLocalName.length + localExtra.length + compressed.length + descriptor.length;

    centralChunks.push(
      buildCentralHeader({
        flags,
        method,
        crc,
        compressedSize,
        fileSize,
        nameLength: rawCentralName.length,
        extraLength: centralExtra.length,
        commentLength: comment.length,
        diskStart: spec.diskStart ?? 0,
        headerOffset,
      }),
      rawCentralName,
      centralExtra,
      comment,
    );
  }

  const centralOffset = bodyLength;
  const centralDirectory = concatBytes(...centralChunks);
  const archiveComment = options.archiveComment ?? new Uint8Array();
  const endRecord = buildEndRecord({
    diskNumber: options.diskNumber ?? 0,
    centralDisk: options.centralDisk ?? 0,
    diskEntries: options.diskEntries ?? members.length,
    totalEntries: options.totalEntries ?? members.length,
    centralSize: centralDirectory.length + (options.centralSizeDelta ?? 0),
    centralOffset,
    commentLength: archiveComment.length,
  });

  return concatBytes(...bodyChunks, centralDirectory, endRecord, archiveComment);
}

function describeZip(bytes: Uint8Array): ZipDescription {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const endOffset = findEndRecord(bytes, view);
  const totalEntries = view.getUint16(endOffset + 10, true);
  const centralSize = view.getUint32(endOffset + 12, true);
  const centralOffset = view.getUint32(endOffset + 16, true);
  const central: ZipCentralRecord[] = [];
  const local: ZipLocalRecord[] = [];
  let cursor = centralOffset;

  for (let index = 0; index < totalEntries; index += 1) {
    assert.equal(view.getUint32(cursor, true), CENTRAL_SIGNATURE, "Invalid central ZIP signature");
    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const headerOffset = view.getUint32(cursor + 42, true);
    const nameStart = cursor + 46;
    const name = decoder.decode(bytes.subarray(nameStart, nameStart + nameLength));
    central.push({ name, flags, method, headerOffset });

    assert.equal(view.getUint32(headerOffset, true), LOCAL_SIGNATURE, `Invalid local ZIP signature for ${name}`);
    const localFlags = view.getUint16(headerOffset + 6, true);
    const localMethod = view.getUint16(headerOffset + 8, true);
    const localNameLength = view.getUint16(headerOffset + 26, true);
    const localName = decoder.decode(bytes.subarray(headerOffset + 30, headerOffset + 30 + localNameLength));
    local.push({ name: localName, flags: localFlags, method: localMethod });

    cursor += 46 + nameLength + extraLength + commentLength;
  }

  assert.equal(cursor, centralOffset + centralSize, "Central directory size mismatch");
  return { central, local, totalEntries, centralSize, centralOffset };
}

function findEndRecord(bytes: Uint8Array, view: DataView): number {
  for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - (22 + 0xffff)); offset -= 1) {
    if (view.getUint32(offset, true) !== END_SIGNATURE) {
      continue;
    }
    const commentLength = view.getUint16(offset + 20, true);
    if (offset + 22 + commentLength === bytes.length) {
      return offset;
    }
  }
  throw new Error("Could not locate ZIP end-of-central-directory record");
}

function buildLocalHeader(fields: {
  flags: number;
  method: number;
  crc: number;
  compressedSize: number;
  fileSize: number;
  nameLength: number;
  extraLength: number;
}): Uint8Array {
  const bytes = new Uint8Array(30);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, LOCAL_SIGNATURE, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, fields.flags, true);
  view.setUint16(8, fields.method, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 33, true);
  view.setUint32(14, fields.crc >>> 0, true);
  view.setUint32(18, fields.compressedSize >>> 0, true);
  view.setUint32(22, fields.fileSize >>> 0, true);
  view.setUint16(26, fields.nameLength, true);
  view.setUint16(28, fields.extraLength, true);
  return bytes;
}

function buildCentralHeader(fields: {
  flags: number;
  method: number;
  crc: number;
  compressedSize: number;
  fileSize: number;
  nameLength: number;
  extraLength: number;
  commentLength: number;
  diskStart: number;
  headerOffset: number;
}): Uint8Array {
  const bytes = new Uint8Array(46);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, CENTRAL_SIGNATURE, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, fields.flags, true);
  view.setUint16(10, fields.method, true);
  view.setUint16(12, 0, true);
  view.setUint16(14, 33, true);
  view.setUint32(16, fields.crc >>> 0, true);
  view.setUint32(20, fields.compressedSize >>> 0, true);
  view.setUint32(24, fields.fileSize >>> 0, true);
  view.setUint16(28, fields.nameLength, true);
  view.setUint16(30, fields.extraLength, true);
  view.setUint16(32, fields.commentLength, true);
  view.setUint16(34, fields.diskStart, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, fields.headerOffset >>> 0, true);
  return bytes;
}

function buildEndRecord(fields: {
  diskNumber: number;
  centralDisk: number;
  diskEntries: number;
  totalEntries: number;
  centralSize: number;
  centralOffset: number;
  commentLength: number;
}): Uint8Array {
  const bytes = new Uint8Array(22);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, END_SIGNATURE, true);
  view.setUint16(4, fields.diskNumber, true);
  view.setUint16(6, fields.centralDisk, true);
  view.setUint16(8, fields.diskEntries, true);
  view.setUint16(10, fields.totalEntries, true);
  view.setUint32(12, fields.centralSize >>> 0, true);
  view.setUint32(16, fields.centralOffset >>> 0, true);
  view.setUint16(20, fields.commentLength, true);
  return bytes;
}

function buildDataDescriptor(fields: {
  crc: number;
  compressedSize: number;
  fileSize: number;
  withSignature: boolean;
}): Uint8Array {
  const bytes = new Uint8Array(fields.withSignature ? 16 : 12);
  const view = new DataView(bytes.buffer);
  let offset = 0;
  if (fields.withSignature) {
    view.setUint32(0, DATA_DESCRIPTOR_SIGNATURE, true);
    offset = 4;
  }
  view.setUint32(offset, fields.crc >>> 0, true);
  view.setUint32(offset + 4, fields.compressedSize >>> 0, true);
  view.setUint32(offset + 8, fields.fileSize >>> 0, true);
  return bytes;
}

function compress(method: number, blob: Uint8Array): Uint8Array {
  if (method === METHOD_STORED) {
    return blob;
  }
  if (method === METHOD_DEFLATED) {
    return new Uint8Array(deflateRawSync(Buffer.from(blob)));
  }
  return blob;
}

function concatBytes(...chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const merged = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }
  return merged;
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  return Buffer.compare(Buffer.from(left), Buffer.from(right)) === 0;
}

function assertBytesEqual(left: Uint8Array, right: Uint8Array, label: string): void {
  assert.equal(Buffer.compare(Buffer.from(left), Buffer.from(right)), 0, `${label} bytes differ`);
}
