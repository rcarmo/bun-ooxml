import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import { afterEach, describe, expect, test } from "bun:test";
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { Presentation } from "../../src/pptx/index.ts";
import {
  bindings,
  runAtomicRefusalScenario,
  runFormattingPreservationScenario,
  runGeometryRoundTripScenario,
  runStaleTableHandleScenario,
} from "../acceptance/tables-pptx.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const roots: string[] = [];

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("pptx Presentation tables", () => {
  test("round-trips saved table geometry sums and text after reopen", async () => {
    const result = await runGeometryRoundTripScenario();

    expect(result.reopenedRows).toBe(2);
    expect(result.reopenedColumns).toBe(3);
    expect(result.reopenedText).toBe("  <Alpha & Beta>  ");
    expect(result.metrics).toEqual({
      x: 120,
      y: 240,
      width: 1001,
      height: 1003,
      columnWidths: [333, 333, 335],
      rowHeights: [501, 502],
    });

    const reopened = await Presentation.open(result.savedBytes);
    expect(reopened.slides[0]?.tables[0]?.cell(0, 0).text).toBe("  <Alpha & Beta>  ");
  });

  test("preserves simple table cell formatting when replacing text and reopening", async () => {
    const result = await runFormattingPreservationScenario();

    expect(result.reopenedText).toBe("updated value");
    expect(result.formatting).toEqual({
      bodyPrAttrs: { wrap: "square" },
      tcPrAttrs: { marL: "111" },
      fillColor: "FFFF00",
      paragraphAttrs: { algn: "r" },
      firstRunAttrs: { b: "1", sz: "1800" },
      endParaRPrAttrs: { lang: "en-US" },
    });
  });

  test("stale table cell handles refuse atomically after slide mutation", async () => {
    const result = await runStaleTableHandleScenario();

    expect(result.error.code).toBe("PPTX_STALE_TABLE_HANDLE");
    expect(result.bytesAfterRefusal).toEqual(result.bytesAfterMutation);
  });

  test("refuses merged and malformed table topologies atomically", async () => {
    const merged = await runAtomicRefusalScenario("merged-cell");
    expect(merged.error.code).toBe("PPTX_TABLE_MERGE_UNSUPPORTED");
    expect(merged.bytesAfterRefusal).toEqual(merged.sourceBytes);
    expect(merged.diffAfterRefusal).toEqual({ added: [], changed: [], removed: [] });

    const malformed = await runAtomicRefusalScenario("malformed-merge");
    expect(malformed.error.code).toBe("PPTX_TABLE_STRUCTURE_UNSUPPORTED");
    expect(malformed.bytesAfterRefusal).toEqual(malformed.sourceBytes);
    expect(malformed.diffAfterRefusal).toEqual({ added: [], changed: [], removed: [] });
  });

  test("passes the PPTX table acceptance feature with dedicated bindings", async () => {
    const root = await makeAcceptanceRoot();
    const report = await runAcceptance(bindings, { root });

    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.inventory.scenarios.implemented).toBe(4);
    expect(report.inventory.cases.implemented).toBe(5);
    expect(report.execution.cases.passed).toBe(5);
    expect(report.execution.steps.failed).toBe(0);
    expect(report.failures).toEqual([]);
  }, 300_000);
});

async function makeAcceptanceRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-pptx-table-acceptance-"));
  roots.push(root);
  const featureText = (await Bun.file(join(fixturesRoot(), "workflows/pptx/tables.feature")).text()).replace(/^@planned/m, '@implemented @bun');
  const path = join(root, "features", "pptx", "tables.feature");
  await mkdir(dirname(path), { recursive: true });
  await Bun.write(path, featureText);
  return root;
}
