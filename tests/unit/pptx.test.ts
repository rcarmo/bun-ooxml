import { afterEach, describe, expect, test } from "bun:test";
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { readZip } from "../../src/opc/zip.ts";
import { Presentation } from "../../src/pptx/index.ts";
import { fixturePath, fixturesRoot, F } from "../../scripts/fixture-inputs.ts";
import {
  bindings,
  runCrossRunReplacementScenario,
  runOrderedNotesScenario,
  runReadableUnsupportedScenario,
  runStaleAnchorScenario,
} from "../acceptance/pptx.ts";

const roots: string[] = [];
const encoder = new TextEncoder();

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("pptx slice", () => {
  test("acceptance feature passes with the dedicated PPTX bindings", async () => {
    const root = await makeProject({
      "features/pptx/text.feature": (await Bun.file(join(fixturesRoot(), "workflows/native/pptx-text.feature")).text()).replace(/^@planned/m, "@implemented @bun"),
    });

    const report = await runAcceptance(bindings, { root });
    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.execution.cases.passed).toBe(4);
  });

  test("follows presentation relationships, preserves notes blank lines and visible fields, and keeps notes reads read-only", async () => {
    const result = await runOrderedNotesScenario();
    expect(result.titles).toEqual([
      "Chapter 5",
      "Frankenstein Lecture Series",
      "Chapter 4",
      "Chapter 2",
      "Chapter 3",
    ]);
    expect(result.notes).toEqual([
      "Key themes:\ncreation, responsibility, isolation\n\nVisible date: 2026-03-12",
      "Compare to Prometheus myth",
    ]);
    expect(result.orderedDiff).toEqual({
      added: [],
      changed: [],
      removed: [],
    });
    expect(result.orderedBytesAfterRead).toEqual(result.orderedBytes);
    expect(result.minimalError.code).toBe("PPTX_NOTES_MISSING");
    expect(result.minimalBytesAfterRefusal).toEqual(result.minimalSource);
    expect(result.minimalNames.some((name) => name.startsWith("ppt/notesSlides/"))).toBeFalse();
  });

  test("exposes line breaks and field text in inspectText and refuses editing that topology atomically", async () => {
    const result = await runReadableUnsupportedScenario();
    expect(result.paragraphText).toBe("Chapter\n7 Notes");
    expect(result.paragraphRuns).toEqual([
      { text: "Chapter", attrs: { b: "1" } },
      { text: "\n", attrs: { lang: "en-US" } },
      { text: "7", attrs: { i: "1" } },
      { text: " Notes", attrs: { u: "sng" } },
    ]);
    expect(result.error.code).toBe("PPTX_UNSUPPORTED_TEXT_TOPOLOGY");
    expect(result.diffAfterRefusal).toEqual({
      added: [],
      changed: [],
      removed: [],
    });
    expect(result.bytesAfterRefusal).toEqual(result.sourceBytes);
  });

  test("replaces exact anchored text across runs and preserves untouched members", async () => {
    const result = await runCrossRunReplacementScenario();
    expect(result.reopenedTitle).toBe("Friendstein");
    expect(result.changedParts).toEqual({
      added: [],
      changed: ["ppt/slides/slide1.xml"],
      removed: [],
    });
    expect(result.customMemberBytes).toEqual(encoder.encode("keep-me-safe"));
    const expectedDocProps = readZip(result.sourceBytes).get("docProps/app.xml");
    if (!expectedDocProps) {
      throw new Error("expected docProps/app.xml in source fixture");
    }
    expect(result.docPropsBytes).toEqual(expectedDocProps);
    expect(result.titleRuns).toEqual([
      { text: "Fr", attrs: { b: "1" } },
      { text: "iend", attrs: { b: "1" } },
      { text: "stein", attrs: { u: "sng" } },
    ]);
  });

  test("refuses stale anchors atomically", async () => {
    const result = await runStaleAnchorScenario();
    expect(result.error.code).toBe("PPTX_STALE_ANCHOR");
    expect(result.reopenedTitle).toBe("Creature");
    expect(result.bytesAfterRefusal).toEqual(result.bytesAfterSuccess);
  });

  test("opens from path and bytes and saves a no-op deck byte-identically", async () => {
    const sourcePath = fixturePath(F.officeSlides.titleSlide);
    const sourceBytes = Uint8Array.from(await Bun.file(sourcePath).bytes());

    const fromPath = await Presentation.open(sourcePath);
    const fromBytes = await Presentation.open(sourceBytes);

    expect(fromPath.slides[0]?.inspectText("unit.open.path")[0]?.text).toBe("Frankenstein");
    expect(fromBytes.package.toBytes()).toEqual(sourceBytes);

    const root = await mkdtemp(join(tmpdir(), "bun-ooxml-pptx-open-"));
    roots.push(root);
    const output = join(root, "title-slide-copy.pptx");
    await fromPath.save(output);
    expect(Uint8Array.from(await Bun.file(output).bytes())).toEqual(sourceBytes);
  });
});

async function makeProject(files: Record<string, string>): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-pptx-gherkin-"));
  roots.push(root);
  for (const [relativePath, content] of Object.entries(files)) {
    const path = join(root, relativePath);
    await mkdir(dirname(path), { recursive: true });
    await Bun.write(path, content);
  }
  return root;
}
