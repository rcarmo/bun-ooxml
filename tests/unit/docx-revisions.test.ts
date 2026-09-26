import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import { afterEach, describe, expect, test } from "bun:test";
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { inspectRevisions, resolveRevisions } from "../../src/docx/revisions.ts";
import { OoxmlError } from "../../src/errors.ts";
import { bindings, openRevisionFixture, readRevisionStoryText } from "../acceptance/revisions-docx.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const tempRoots: string[] = [];

afterEach(async () => {
  await Promise.all(tempRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("docx revisions", () => {
  test("discovers reachable nested story parts for inspection and selected resolution", async () => {
    const pkg = await openRevisionFixture("all-stories");

    expect([...new Set(inspectRevisions(pkg).revisions.map((revision) => revision.part))]).toEqual([
      "word/document.xml",
      "word/header1.xml",
      "word/footer1.xml",
      "word/footer2.xml",
      "word/footnotes.xml",
      "word/endnotes.xml",
      "word/comments.xml",
    ]);

    expect(resolveRevisions(pkg, "accept", { parts: ["word/footer2.xml"] })).toEqual({
      resolved: 2,
      changedParts: ["word/footer2.xml"],
    });
    expect(await readRevisionStoryText(pkg, "word/footer2.xml", "all")).toBe("Nested footer modernized.");
    expect(inspectRevisions(pkg).revisions.filter((revision) => revision.part === "word/footer2.xml")).toEqual([]);
  });

  test("preserves namespaced run properties when unwrapping and renames delText on reject", async () => {
    const accepted = await openRevisionFixture("all-stories");
    const rejected = await openRevisionFixture("all-stories");

    resolveRevisions(accepted, "accept", { parts: ["word/footer2.xml"] });
    resolveRevisions(rejected, "reject", { parts: ["word/footer2.xml"] });

    const acceptedXml = accepted.text("word/footer2.xml");
    const rejectedXml = rejected.text("word/footer2.xml");

    expect(acceptedXml).toContain(`<w:r xmlns:w14="${W14_NS}"><w:rPr><w14:glow w14:rad="50000"/></w:rPr><w:t>Nested footer modernized.</w:t></w:r>`);
    expect(acceptedXml).not.toContain("<w:ins");
    expect(acceptedXml).not.toContain("<w:delText>");

    expect(rejectedXml).toContain(`<w:r xmlns:w14="${W14_NS}"><w:rPr><w14:glow w14:rad="50000"/></w:rPr><w:t>Nested footer historic.</w:t></w:r>`);
    expect(rejectedXml).not.toContain("<w:del");
    expect(rejectedXml).not.toContain("w:delText");
  });

  test('invalid revision IDs and lexical text between run children refuse atomically',async()=>{
    for(const kind of ['id','text']) {
      const p=await openRevisionFixture('all-stories');let xml=p.text('word/document.xml');
      xml=kind==='id'?xml.replace(/w:id="[^"]+"/,'w:id="not-an-id"'):xml.replace(/(<w:ins\b[^>]*>)/,'$1UNOWNED');
      p.set('word/document.xml',xml);const before=p.toBytes();
      expect(()=>resolveRevisions(p,'accept')).toThrow();expect(p.toBytes()).toEqual(before);
    }
  });
  test("rejects unsafe namespace lifts atomically", async () => {
    const pkg = await openRevisionFixture("unsafe-lift");
    const before = pkg.toBytes();

    let refusal: unknown;
    try {
      resolveRevisions(pkg, "reject");
    } catch (error) {
      refusal = error;
    }

    expect(refusal).toBeInstanceOf(OoxmlError);
    expect((refusal as OoxmlError).code).toBe("docx-revisions-unsupported");
    expect((refusal as OoxmlError).message).toContain("cannot safely lift xmlns:w14");
    expect(pkg.toBytes()).toEqual(before);
  });

  test("preflights unsupported nested stories and protected packages before mutating bytes", async () => {
    const unsupported = await openRevisionFixture("rollback-unsupported");
    const unsupportedBefore = unsupported.toBytes();
    const inspected = inspectRevisions(unsupported);

    expect(inspected.revisions).toHaveLength(2);
    expect(inspected.unsupported).toContainEqual({
      part: "word/footer2.xml",
      kind: "format",
      reason: "format revisions are not supported",
    });
    expect(() => resolveRevisions(unsupported, "accept", { parts: ["word/document.xml", "word/footer2.xml"] })).toThrow(
      expect.objectContaining({ code: "docx-revisions-unsupported" }),
    );
    expect(unsupported.toBytes()).toEqual(unsupportedBefore);

    const protectedPkg = await openRevisionFixture("protected");
    const protectedBefore = protectedPkg.toBytes();
    expect(() => resolveRevisions(protectedPkg, "accept")).toThrow(
      expect.objectContaining({ code: "docx-revisions-protected" }),
    );
    expect(protectedPkg.toBytes()).toEqual(protectedBefore);
  });

  test('missing enforcement/edit and duplicate protection declarations cannot bypass refusal', async () => {
    for (const replacement of [
      '<w:documentProtection w:enforcement="1"/>',
      '<w:documentProtection w:edit="readOnly"/>',
      '<w:documentProtection w:enforcement="0"/><w:documentProtection w:enforcement="1" w:edit="readOnly"/>',
    ]) {
      const pkg = await openRevisionFixture('protected');
      const settings = pkg.related(pkg.mainPart(), 'settings')!;
      const beforeXml = pkg.text(settings);
      const changed = beforeXml.replace(/<w:documentProtection\b[^>]*\/>/, replacement);
      expect(changed).not.toBe(beforeXml);
      pkg.set(settings, changed); const before = pkg.toBytes();
      expect(() => resolveRevisions(pkg, 'accept')).toThrow(expect.objectContaining({code: 'docx-revisions-protected'}));
      expect(pkg.toBytes()).toEqual(before);
    }
  });

  test("passes the DOCX revisions acceptance feature with dedicated bindings", async () => {
    const root = await makeAcceptanceRoot();
    const report = await runAcceptance(bindings, { root });

    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.inventory.scenarios.implemented).toBe(4);
    expect(report.execution.cases.passed).toBe(4);
    expect(report.execution.steps.failed).toBe(0);
    expect(report.failures).toEqual([]);
  });
});

async function makeAcceptanceRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-docx-revisions-acceptance-"));
  tempRoots.push(root);
  const featureText = (await Bun.file(join(fixturesRoot(), "workflows/native/docx-revisions.feature")).text()).replace(/^@planned/m, '@implemented @bun');
  const path = join(root, "features", "docx", "revisions.feature");
  await mkdir(dirname(path), { recursive: true });
  await Bun.write(path, featureText);
  return root;
}
