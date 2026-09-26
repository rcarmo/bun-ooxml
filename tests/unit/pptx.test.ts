import { afterEach, describe, expect, test } from "bun:test";
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { Presentation } from "../../src/pptx/index.ts";
import {
  bindings,
  runCrossRunReplacementScenario,
  runOrderedNotesScenario,
  runStaleAnchorScenario,
} from "../acceptance/pptx.ts";

const roots: string[] = [];

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("pptx slice", () => {
  test("acceptance feature passes with the dedicated PPTX bindings", async () => {
    const root = await makeProject({
      "features/pptx/text.feature": await Bun.file(
        join(import.meta.dir, "../../features/pptx/text.feature"),
      ).text(),
    });

    const report = await runAcceptance(bindings, { root });
    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.execution.cases.passed).toBe(3);
  });

  test("follows presentation relationships and keeps notes reads read-only", async () => {
    const result = await runOrderedNotesScenario();
    expect(result.titles[0]).toBe("Chapter 5");
    expect(result.notes[0]).toContain("creation");
    expect(result.minimalError.code).toBe("PPTX_NOTES_MISSING");
  });

  test("replaces exact anchored text across runs and preserves untouched members", async () => {
    const result = await runCrossRunReplacementScenario();
    expect(result.reopenedTitle).toBe("Friendstein");
    expect(result.changedParts.changed).toEqual(["ppt/slides/slide1.xml"]);
    expect(result.titleRuns.map((run) => run.text)).toEqual(["Fr", "iend", "stein"]);
  });

  test("refuses stale anchors atomically", async () => {
    const result = await runStaleAnchorScenario();
    expect(result.error.code).toBe("PPTX_STALE_ANCHOR");
    expect(result.reopenedTitle).toBe("Creature");
    expect(result.bytesAfterRefusal).toEqual(result.bytesAfterSuccess);
  });

  test("opens from path and bytes and saves a no-op deck byte-identically", async () => {
    const sourcePath = join(
      import.meta.dir,
      "../../fixtures/python-office-mcp-server/tests/_templates/testdata/pptx/title_slide.pptx",
    );
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
