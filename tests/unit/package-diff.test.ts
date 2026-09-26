import { describe, expect, test } from "bun:test";
import { join, resolve } from "node:path";

import { diffPackages } from "../../src/opc/diff.ts";
import { OpcPackage } from "../../src/opc/package.ts";
import { writeZip } from "../../src/opc/zip.ts";

const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const MAIN_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const FIXTURE_PATH = join(resolve(import.meta.dir, "../.."), "docs/contracts/shared-v2/pack/fixtures/present-placeholder.docx");
const encoder = new TextEncoder();


describe("diffPackages", () => {
  test("raw fixture no-op diffs are empty, preserve exact source bytes, and return detached reports", async () => {
    const original = Uint8Array.from(await Bun.file(FIXTURE_PATH).bytes());
    const before = await OpcPackage.open(original);
    const after = await OpcPackage.open(original.slice());

    const report = diffPackages(before, after);

    expect(report).toEqual({
      added: [],
      removed: [],
      changed: [],
      unchanged: before.names(),
      parts: [],
    });

    report.unchanged.pop();
    report.parts.push({ name: "invented", kind: "added" });

    expect(before.toBytes()).toEqual(original);
    expect(after.toBytes()).toEqual(original);
    expect(diffPackages(before, after)).toEqual({
      added: [],
      removed: [],
      changed: [],
      unchanged: before.names(),
      parts: [],
    });
  });

  test("ignores ZIP metadata-only changes", async () => {
    const bytes = createPackageBytes();
    const metadataOnly = withZipComment(bytes, "metadata-only");
    const before = await OpcPackage.open(bytes);
    const after = await OpcPackage.open(metadataOnly);

    expect(metadataOnly).not.toEqual(bytes);
    expect(diffPackages(before, after)).toEqual({
      added: [],
      removed: [],
      changed: [],
      unchanged: before.names(),
      parts: [],
    });
    expect(before.toBytes()).toEqual(bytes);
    expect(after.toBytes()).toEqual(metadataOnly);
  });

  test("reports opaque payload edits with hashes, lengths, content types, and detached part records", async () => {
    const beforeBytes = createPackageBytes({
      parts: { "custom/data.bin": Uint8Array.of(1, 2, 3) },
      overrides: { "custom/data.bin": "application/vnd.example.opaque" },
    });
    const before = await OpcPackage.open(beforeBytes);
    const after = await OpcPackage.open(beforeBytes);
    const replacement = Uint8Array.of(3, 2, 1, 0);

    after.set("custom/data.bin", replacement);

    const report = diffPackages(before, after);

    expect(report.added).toEqual([]);
    expect(report.removed).toEqual([]);
    expect(report.changed).toEqual(["custom/data.bin"]);
    expect(report.unchanged).toEqual(["[Content_Types].xml", "_rels/.rels", "word/document.xml"]);
    expect(report.parts).toEqual([
      {
        name: "custom/data.bin",
        kind: "changed",
        beforeSha256: sha(Uint8Array.of(1, 2, 3)),
        afterSha256: sha(replacement),
        beforeBytes: 3,
        afterBytes: 4,
        beforeContentType: "application/vnd.example.opaque",
        afterContentType: "application/vnd.example.opaque",
      },
    ]);

    report.changed[0] = "mutated";
    report.parts[0]!.afterSha256 = "not-the-real-hash";

    expect(after.get("custom/data.bin")).toEqual(replacement);
    expect(before.toBytes()).toEqual(beforeBytes);
    expect(diffPackages(before, after).parts).toEqual([
      {
        name: "custom/data.bin",
        kind: "changed",
        beforeSha256: sha(Uint8Array.of(1, 2, 3)),
        afterSha256: sha(replacement),
        beforeBytes: 3,
        afterBytes: 4,
        beforeContentType: "application/vnd.example.opaque",
        afterContentType: "application/vnd.example.opaque",
      },
    ]);
  });

  test("reports added and removed parts with content-type updates via package.set", async () => {
    const beforeBytes = createPackageBytes({
      parts: {
        "custom/existing.bin": Uint8Array.of(9, 9),
        "custom/obsolete.xml": xml("<obsolete>before</obsolete>"),
      },
      overrides: {
        "custom/existing.bin": "application/vnd.example.existing",
        "custom/obsolete.xml": "application/vnd.example.obsolete+xml",
      },
    });
    const before = await OpcPackage.open(beforeBytes);
    const after = await OpcPackage.open(beforeBytes);
    const afterPayload = xml("<fresh>after</fresh>");

    after.delete("custom/obsolete.xml");
    after.set("custom/fresh.xml", afterPayload);
    after.set(
      "[Content_Types].xml",
      after.text("[Content_Types].xml")
        .replace('<Override PartName="/custom/obsolete.xml" ContentType="application/vnd.example.obsolete+xml"/>', "")
        .replace("</Types>", '<Override PartName="/custom/fresh.xml" ContentType="application/vnd.example.fresh+xml"/></Types>'),
    );

    const report = diffPackages(before, after);

    expect(report.added).toEqual(["custom/fresh.xml"]);
    expect(report.removed).toEqual(["custom/obsolete.xml"]);
    expect(report.changed).toEqual(["[Content_Types].xml"]);
    expect(report.unchanged).toEqual(["_rels/.rels", "custom/existing.bin", "word/document.xml"]);
    expect(report.parts).toEqual([
      {
        name: "[Content_Types].xml",
        kind: "changed",
        beforeSha256: sha(before.get("[Content_Types].xml")!),
        afterSha256: sha(after.get("[Content_Types].xml")!),
        beforeBytes: before.get("[Content_Types].xml")!.length,
        afterBytes: after.get("[Content_Types].xml")!.length,
        beforeContentType: "application/xml",
        afterContentType: "application/xml",
      },
      {
        name: "custom/fresh.xml",
        kind: "added",
        afterSha256: sha(afterPayload),
        afterBytes: afterPayload.length,
        afterContentType: "application/vnd.example.fresh+xml",
      },
      {
        name: "custom/obsolete.xml",
        kind: "removed",
        beforeSha256: sha(before.get("custom/obsolete.xml")!),
        beforeBytes: before.get("custom/obsolete.xml")!.length,
        beforeContentType: "application/vnd.example.obsolete+xml",
      },
    ]);
  });

  test("marks content-type-only changes as changed even when payload bytes are identical", async () => {
    const payload = Uint8Array.of(7, 8, 9);
    const beforeBytes = createPackageBytes({
      parts: { "custom/payload.bin": payload },
      overrides: { "custom/payload.bin": "application/vnd.example.before" },
    });
    const before = await OpcPackage.open(beforeBytes);
    const after = await OpcPackage.open(beforeBytes);

    after.set(
      "[Content_Types].xml",
      after.text("[Content_Types].xml").replace("application/vnd.example.before", "application/vnd.example.after"),
    );

    const report = diffPackages(before, after);

    expect(report.added).toEqual([]);
    expect(report.removed).toEqual([]);
    expect(report.changed).toEqual(["[Content_Types].xml", "custom/payload.bin"]);
    expect(report.unchanged).toEqual(["_rels/.rels", "word/document.xml"]);
    expect(report.parts).toEqual([
      {
        name: "[Content_Types].xml",
        kind: "changed",
        beforeSha256: sha(before.get("[Content_Types].xml")!),
        afterSha256: sha(after.get("[Content_Types].xml")!),
        beforeBytes: before.get("[Content_Types].xml")!.length,
        afterBytes: after.get("[Content_Types].xml")!.length,
        beforeContentType: "application/xml",
        afterContentType: "application/xml",
      },
      {
        name: "custom/payload.bin",
        kind: "changed",
        beforeSha256: sha(payload),
        afterSha256: sha(payload),
        beforeBytes: payload.length,
        afterBytes: payload.length,
        beforeContentType: "application/vnd.example.before",
        afterContentType: "application/vnd.example.after",
      },
    ]);
  });
});

function createPackageBytes(options: {
  parts?: Record<string, Uint8Array | string>;
  overrides?: Record<string, string>;
  defaults?: Record<string, string>;
} = {}): Uint8Array {
  const defaults = {
    rels: RELATIONSHIPS_CONTENT_TYPE,
    xml: "application/xml",
    ...(options.defaults ?? {}),
  };
  const overrides = {
    "word/document.xml": MAIN_CONTENT_TYPE,
    ...(options.overrides ?? {}),
  };
  const parts = new Map<string, Uint8Array>([
    ["[Content_Types].xml", encoder.encode(contentTypesXml(defaults, overrides))],
    ["_rels/.rels", encoder.encode(relationshipsXml("word/document.xml"))],
    ["word/document.xml", xml("<document>Main</document>")],
  ]);

  for (const [name, value] of Object.entries(options.parts ?? {})) {
    parts.set(name, typeof value === "string" ? encoder.encode(value) : value.slice());
  }

  return writeZip(parts);
}

function contentTypesXml(defaults: Record<string, string>, overrides: Record<string, string>): string {
  return `<?xml version="1.0" encoding="UTF-8"?>`
    + `<Types xmlns="${CONTENT_TYPES_NS}">`
    + Object.keys(defaults).sort().map((extension) => `<Default Extension="${extension}" ContentType="${defaults[extension]}"/>`).join("")
    + Object.keys(overrides).sort().map((name) => `<Override PartName="/${name}" ContentType="${overrides[name]}"/>`).join("")
    + `</Types>`;
}

function relationshipsXml(target: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>`
    + `<Relationships xmlns="${RELATIONSHIPS_NS}">`
    + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="${target}"/>`
    + `</Relationships>`;
}

function xml(body: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>${body}`);
}

function sha(bytes: Uint8Array): string {
  return new Bun.CryptoHasher("sha256").update(bytes).digest("hex");
}

function withZipComment(bytes: Uint8Array, comment: string): Uint8Array {
  const signature = [0x50, 0x4b, 0x05, 0x06];
  let offset = -1;
  for (let index = bytes.length - 22; index >= 0; index -= 1) {
    if (
      bytes[index] === signature[0]
      && bytes[index + 1] === signature[1]
      && bytes[index + 2] === signature[2]
      && bytes[index + 3] === signature[3]
    ) {
      offset = index;
      break;
    }
  }
  if (offset === -1) throw new Error("ZIP end-of-central-directory not found");

  const commentBytes = encoder.encode(comment);
  const next = new Uint8Array(bytes.length + commentBytes.length);
  next.set(bytes);
  new DataView(next.buffer).setUint16(offset + 20, commentBytes.length, true);
  next.set(commentBytes, bytes.length);
  return next;
}
