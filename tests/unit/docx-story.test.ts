import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import { afterEach, describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import {textboxDocument} from '../fixtures/native-edge-cases.ts';
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { inspectStories, storyParts } from "../../src/docx/story.ts";
import { OpcPackage } from "../../src/opc/index.ts";
import {
  bindings as storyBindings,
  createSyntheticBlindFixture,
  createSyntheticRevisionFixture,
  resolveStoryFixturePath,
} from "../acceptance/story-docx.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const tempRoots: string[] = [];

afterEach(async () => {
  await Promise.all(tempRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("docx story inspection", () => {
  test("lists linked native story parts in deterministic order", async () => {
    const pkg = await createSyntheticRevisionFixture();

    expect(storyParts(pkg)).toEqual([
      { part: "word/document.xml", kind: "body" },
      { part: "word/header1.xml", kind: "header" },
      { part: "word/footer1.xml", kind: "footer" },
      { part: "word/footer2.xml", kind: "footer" },
      { part: "word/footnotes.xml", kind: "footnotes" },
      { part: "word/endnotes.xml", kind: "endnotes" },
      { part: "word/comments.xml", kind: "comments" },
    ]);
  });

  test("reads owned document headers/comments and native notes without inferring from filenames", async () => {
    const headerPkg = await OpcPackage.open(resolveStoryFixturePath("header-footer-sections"));
    const notePkg = await createSyntheticRevisionFixture();
    const commentPkg = await OpcPackage.open(resolveStoryFixturePath("comments"));

    const headers = inspectStories(headerPkg);
    const notes = inspectStories(notePkg);
    const comments = inspectStories(commentPkg);

    expect(storyText(headers, "word/header1.xml")).toEqual(["FRANKENSTEIN - by Mary Shelley"]);
    expect(storyText(headers, "word/header2.xml")).toEqual(["FRANKENSTEIN - First Page"]);
    expect(storyText(headers, "word/footer1.xml")).toEqual(["Prepared by Test Author"]);
    expect(storyText(notes, "word/document.xml")).toEqual(["Body intro", "Table\tCell\nLine2", "Inserted body paragraph", "Body kept"]);
    expect(storyText(notes, "word/footnotes.xml")).toEqual([" Footnote new text"]);
    expect(storyText(notes, "word/endnotes.xml")).toEqual([" Endnote text"]);
    expect(storyText(comments, "word/comments.xml")).toEqual([
      "This is a great opening", "Classical hubris", "(this is a threaded reply)",
    ]);
  });

  test("inspects current, original and all views across synthetic body/header/footer/note stories with renamed prefixes", async () => {
    const pkg = await createSyntheticRevisionFixture();

    expect(storyParts(pkg)).toEqual([
      { part: "word/document.xml", kind: "body" },
      { part: "word/header1.xml", kind: "header" },
      { part: "word/footer1.xml", kind: "footer" },
      { part: "word/footer2.xml", kind: "footer" },
      { part: "word/footnotes.xml", kind: "footnotes" },
      { part: "word/endnotes.xml", kind: "endnotes" },
      { part: "word/comments.xml", kind: "comments" },
    ]);

    expect(inspectStories(pkg, { view: "current" }).stories).toEqual([
      story("word/document.xml", "body", [
        [0, "Body intro"],
        [1, "Table\tCell\nLine2"],
        [2, "Inserted body paragraph"],
        [3, "Body kept"],
      ]),
      story("word/header1.xml", "header", [[0, "Header new text"]]),
      story("word/footer1.xml", "footer", [[0, "Footer current"]]),
      story("word/footer2.xml", "footer", [[0, "Footer via header relationship"]]),
      story("word/footnotes.xml", "footnotes", [[0, " Footnote new text"]]),
      story("word/endnotes.xml", "endnotes", [[0, " Endnote text"]]),
      story("word/comments.xml", "comments", [[0, "Comment current"]]),
    ]);

    expect(inspectStories(pkg, { view: "original" }).stories).toEqual([
      story("word/document.xml", "body", [
        [0, "Body intro"],
        [1, "Table\tCell\nLine2"],
        [2, "Deleted body paragraph"],
        [3, "Body kept"],
      ]),
      story("word/header1.xml", "header", [[0, "Header old text"]]),
      story("word/footer1.xml", "footer", [[0, "Footer original"]]),
      story("word/footer2.xml", "footer", [[0, "Footer via header relationship"]]),
      story("word/footnotes.xml", "footnotes", [[0, " Footnote old text"]]),
      story("word/endnotes.xml", "endnotes", [[0, " Endnote text"]]),
      story("word/comments.xml", "comments", [[0, "Comment original"]]),
    ]);

    expect(inspectStories(pkg, { view: "all" }).stories).toEqual([
      story("word/document.xml", "body", [
        [0, "Body intro"],
        [1, "Table\tCell\nLine2"],
        [2, "Inserted body paragraph"],
        [3, "Deleted body paragraph"],
        [4, "Body kept"],
      ]),
      story("word/header1.xml", "header", [[0, "Header newold text"]]),
      story("word/footer1.xml", "footer", [[0, "Footer currentoriginal"]]),
      story("word/footer2.xml", "footer", [[0, "Footer via header relationship"]]),
      story("word/footnotes.xml", "footnotes", [[0, " Footnote newold text"]]),
      story("word/endnotes.xml", "endnotes", [[0, " Endnote text"]]),
      story("word/comments.xml", "comments", [[0, "Comment currentoriginal"]]),
    ]);
  });

  test("counts blind regions for native text boxes fields alternate content unknown namespaces orphans and external story rels", async () => {
    const textboxPkg = await OpcPackage.open(await textboxDocument());
    const blindPkg = await createSyntheticBlindFixture();

    const textbox = inspectStories(textboxPkg);
    expect(storyText(textbox, "word/document.xml")).not.toContain("Text living inside the text box.");
    expect(textbox.blindRegions).toContainEqual({ part: "word/document.xml", kind: "textbox", count: 1 });

    const blind = inspectStories(blindPkg);
    expect(blind.blindRegions).toEqual([
      { part: "word/document.xml", kind: "alternate-content", count: 1 },
      { part: "word/document.xml", kind: "field", count: 2 },
      { part: "word/document.xml", kind: "textbox", count: 1 },
      { part: "word/document.xml", kind: "unknown-namespace", count: 1 },
      { part: "word/document.xml", kind: "external-story-relationship", count: 1 },
      { part: "word/header9.xml", kind: "orphan-story-part", count: 1 },
    ]);
    expect(storyText(blind, "word/document.xml")).toEqual([
      "Field 2026-01-01",
      "7",
      "Opaque text",
      "Choice text",
      "Outside",
    ]);
  });

  test("inspection never mutates package bytes", async () => {
    const source = new Uint8Array(readFileSync(resolveStoryFixturePath("header-footer-sections")));
    const pkg = await OpcPackage.open(source);

    inspectStories(pkg, { view: "all" });

    expect(pkg.toBytes()).toEqual(source);
  });

  test("passes the DOCX story acceptance feature with dedicated bindings", async () => {
    const root = await makeAcceptanceRoot();
    const report = await runAcceptance(storyBindings, { root });

    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.inventory.scenarios.implemented).toBe(3);
    expect(report.execution.cases.passed).toBe(3);
    expect(report.execution.steps.failed).toBe(0);
    expect(report.failures).toEqual([]);
  });
});

function story(
  part: string,
  kind: "body" | "header" | "footer" | "footnotes" | "endnotes" | "comments",
  paragraphs: Array<[number, string]>,
) {
  return {
    part,
    kind,
    paragraphs: paragraphs.map(([index, text]) => ({ index, text })),
  };
}

function storyText(
  inspection: ReturnType<typeof inspectStories>,
  part: string,
): string[] {
  return inspection.stories.find((story) => story.part === part)?.paragraphs.map((paragraph) => paragraph.text) ?? [];
}

async function makeAcceptanceRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-docx-story-acceptance-"));
  tempRoots.push(root);
  const featureText = (await Bun.file(join(fixturesRoot(), "workflows/native/docx-story.feature")).text()).replace(/^@planned/m, '@implemented @bun');
  const path = join(root, "features", "docx", "story.feature");
  await mkdir(dirname(path), { recursive: true });
  await Bun.write(path, featureText);
  return root;
}
