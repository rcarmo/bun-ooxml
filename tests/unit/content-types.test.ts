import { describe, expect, test } from "bun:test";

import { OoxmlError } from "../../src/errors.ts";
import { getContentType, removePartContentType, setPartContentType } from "../../src/opc/content-types.ts";
import { OpcPackage } from "../../src/opc/package.ts";
import { writeZip } from "../../src/opc/zip.ts";

const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const MAIN_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const DOCUMENT_PART = "word/document.xml";
const encoder = new TextEncoder();

describe("opc content type helpers", () => {
  test("getContentType resolves defaults and overrides across namespace prefixes", async () => {
    const pkg = await OpcPackage.open(createPackageBytes({
      contentTypesXml:
        `<?xml version="1.0" encoding="UTF-8"?>`
        + `<ct:Types xmlns:ct="${CONTENT_TYPES_NS}" xmlns:alt="${CONTENT_TYPES_NS}">`
        + `<alt:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
        + `<ct:Default Extension="xml" ContentType="application/xml"/>`
        + `<alt:Override PartName="/${DOCUMENT_PART}" ContentType="${MAIN_CONTENT_TYPE}"/>`
        + `<ct:Override PartName="/custom/data.xml" ContentType="application/custom+xml"/>`
        + `</ct:Types>`,
    }));

    expect(getContentType(pkg, DOCUMENT_PART)).toBe(MAIN_CONTENT_TYPE);
    expect(getContentType(pkg, "word/other.xml")).toBe("application/xml");
    expect(getContentType(pkg, "custom/data.xml")).toBe("application/custom+xml");
  });

  test("setPartContentType is an exact no-op when a default already yields the requested type", async () => {
    const original = createPackageBytes();
    const pkg = await OpcPackage.open(original);

    setPartContentType(pkg, "word/other.xml", "application/xml");
    removePartContentType(pkg, "word/other.xml");

    expect(pkg.toBytes()).toEqual(original);
  });

  test("adds an override to a self-closing prefixed root without rewriting names", async () => {
    const pkg = await OpcPackage.open(createPackageBytes());
    pkg.set(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8"?><ct:Types xmlns:ct="${CONTENT_TYPES_NS}"/>`,
    );

    setPartContentType(pkg, "custom/data.bin", "application/octet-stream");

    expect(pkg.text("[Content_Types].xml")).toBe(
      `<?xml version="1.0" encoding="UTF-8"?><ct:Types xmlns:ct="${CONTENT_TYPES_NS}"><ct:Override PartName="/custom/data.bin" ContentType="application/octet-stream"/></ct:Types>`,
    );
    expect(getContentType(pkg, "custom/data.bin")).toBe("application/octet-stream");
  });

  test("updates an existing override in place and preserves surrounding raw XML", async () => {
    const pkg = await OpcPackage.open(createPackageBytes({
      contentTypesXml:
        `<?xml version="1.0" encoding="UTF-8"?>`
        + `<ct:Types xmlns:ct="${CONTENT_TYPES_NS}">`
        + `<ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
        + `<ct:Override  PartName = '/${DOCUMENT_PART}'  ContentType = 'application/old+xml' />`
        + `</ct:Types>`,
    }));

    setPartContentType(pkg, DOCUMENT_PART, MAIN_CONTENT_TYPE);

    expect(pkg.text("[Content_Types].xml")).toBe(
      `<?xml version="1.0" encoding="UTF-8"?>`
      + `<ct:Types xmlns:ct="${CONTENT_TYPES_NS}">`
      + `<ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + `<ct:Override  PartName = '/${DOCUMENT_PART}'  ContentType = '${MAIN_CONTENT_TYPE}' />`
      + `</ct:Types>`,
    );
  });

  test("removePartContentType removes only the matching override and never rewrites defaults", async () => {
    const originalXml =
      `<?xml version="1.0" encoding="UTF-8"?>`
      + `<Types xmlns="${CONTENT_TYPES_NS}">`
      + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + `<Default Extension="xml" ContentType="application/xml"/>`
      + `<Override PartName="/${DOCUMENT_PART}" ContentType="${MAIN_CONTENT_TYPE}"/>`
      + `<Override PartName="/custom/data.bin" ContentType="application/octet-stream"/>`
      + `</Types>`;
    const pkg = await OpcPackage.open(createPackageBytes({ contentTypesXml: originalXml }));

    removePartContentType(pkg, "custom/data.bin");

    expect(pkg.text("[Content_Types].xml")).toBe(
      `<?xml version="1.0" encoding="UTF-8"?>`
      + `<Types xmlns="${CONTENT_TYPES_NS}">`
      + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + `<Default Extension="xml" ContentType="application/xml"/>`
      + `<Override PartName="/${DOCUMENT_PART}" ContentType="${MAIN_CONTENT_TYPE}"/>`
      + `</Types>`,
    );
    expect(getContentType(pkg, "custom/data.bin")).toBeUndefined();
    expect(getContentType(pkg, "word/other.xml")).toBe("application/xml");
  });

  test("detached overrides remain queryable for parts that do not yet exist", async () => {
    const pkg = await OpcPackage.open(createPackageBytes());

    setPartContentType(pkg, "custom/future.bin", "application/octet-stream");

    expect(getContentType(pkg, "custom/future.bin")).toBe("application/octet-stream");
    expect(getContentType(pkg, "word/other.xml")).toBe("application/xml");
    expect(() => pkg.toBytes()).not.toThrow();
  });

  test("invalid inputs and malformed current content types leave bytes untouched", async () => {
    const pkg = await OpcPackage.open(createPackageBytes());
    const originalPart = pkg.get("[Content_Types].xml")!;

    expectOoxmlError(
      () => setPartContentType(pkg, "../bad.xml", "application/xml"),
      "opc-part-name-invalid",
      "Noncanonical part name",
    );
    expectOoxmlError(
      () => setPartContentType(pkg, "word/bad.xml", "not a mime"),
      "opc-content-type-invalid",
      "Invalid content type",
    );
    expect(pkg.get("[Content_Types].xml")).toEqual(originalPart);

    pkg.set(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8"?>`
      + `<Types xmlns="${CONTENT_TYPES_NS}">`
      + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + `<Override PartName="/${DOCUMENT_PART}" ContentType="${MAIN_CONTENT_TYPE}"/>`
      + `<Override PartName="/${DOCUMENT_PART}" ContentType="application/duplicate+xml"/>`
      + `</Types>`,
    );
    const duplicateBytes = pkg.get("[Content_Types].xml")!;

    expectOoxmlError(
      () => getContentType(pkg, DOCUMENT_PART),
      "opc-content-types-invalid",
      "Duplicate/missing part override",
    );
    expect(pkg.get("[Content_Types].xml")).toEqual(duplicateBytes);

    pkg.set(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8"?>`
      + `<ct:Types xmlns:ct="${CONTENT_TYPES_NS}">`
      + `<ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + `<bad:Override xmlns:bad="urn:not-content-types" PartName="/${DOCUMENT_PART}" ContentType="${MAIN_CONTENT_TYPE}"/>`
      + `</ct:Types>`,
    );
    const wrongNamespaceBytes = pkg.get("[Content_Types].xml")!;

    expectOoxmlError(
      () => removePartContentType(pkg, DOCUMENT_PART),
      "opc-content-types-invalid",
      "Malformed content type",
    );
    expect(pkg.get("[Content_Types].xml")).toEqual(wrongNamespaceBytes);
  });
});

function createPackageBytes(options: { contentTypesXml?: string } = {}): Uint8Array {
  return writeZip(new Map<string, Uint8Array>([
    ["[Content_Types].xml", encoder.encode(options.contentTypesXml ?? contentTypesXml())],
    ["_rels/.rels", encoder.encode(relationshipsXml())],
    [DOCUMENT_PART, encoder.encode(`<?xml version="1.0" encoding="UTF-8"?><document>Alpha</document>`)],
  ]));
}

function contentTypesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>`
    + `<Types xmlns="${CONTENT_TYPES_NS}">`
    + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
    + `<Default Extension="xml" ContentType="application/xml"/>`
    + `<Override PartName="/${DOCUMENT_PART}" ContentType="${MAIN_CONTENT_TYPE}"/>`
    + `</Types>`;
}

function relationshipsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>`
    + `<Relationships xmlns="${RELATIONSHIPS_NS}">`
    + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="${DOCUMENT_PART}"/>`
    + `</Relationships>`;
}

function expectOoxmlError(action: () => unknown, code: string, fragment: string): void {
  try {
    action();
    throw new Error(`expected ${code}`);
  } catch (error) {
    if (error instanceof Error && error.message === `expected ${code}`) {
      throw error;
    }
    expect(error).toBeInstanceOf(OoxmlError);
    const refusal = error as OoxmlError;
    expect(refusal.code).toBe(code);
    expect(refusal.message).toContain(fragment);
  }
}

