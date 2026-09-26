import { afterEach, describe, expect, test } from "bun:test";
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { PickleStepType } from "@cucumber/messages";
import { runAcceptance } from "../../scripts/acceptance.ts";
import { inventoryFeatures, parseFeature, type StepBinding } from "../../scripts/gherkin.ts";

const roots: string[] = [];

afterEach(async () => {
  await Promise.all(
    roots.splice(0).map((root) => rm(root, { recursive: true, force: true })),
  );
});

describe("gherkin acceptance", () => {
  test("parses backgrounds and scenario outlines into exact expanded cases", () => {
    const feature = parseFeature(
      "features/opc/outline.feature",
      [
        "@implemented @bun",
        "Feature: outline execution",
        "  Background:",
        "    Given a package exists",
        "",
        "  @id-add-part",
        "  Scenario Outline: append <part>",
        "    When I append <part>",
        "    Then the package contains <part>",
        "",
        "    Examples:",
        "      | part      |",
        "      | /a.xml    |",
        "      | /b.xml    |",
        "",
      ].join("\n"),
    );

    expect(feature.lifecycle).toBe("implemented");
    expect(feature.runner).toBe("bun");
    expect(feature.scenarios).toHaveLength(1);
    expect(feature.scenarios[0]?.scenarioId).toBe("@id-add-part");
    expect(feature.scenarios[0]?.cases).toHaveLength(2);

    const firstCase = feature.scenarios[0]?.cases[0];
    expect(firstCase?.caseId).toContain('@id-add-part -- part="/a.xml"');
    expect(firstCase?.steps.map((step) => step.text)).toEqual([
      "a package exists",
      "I append /a.xml",
      "the package contains /a.xml",
    ]);
    expect(firstCase?.steps.map((step) => step.fromBackground)).toEqual([
      true,
      false,
      false,
    ]);
    expect(firstCase?.steps[2]?.type).toBe(PickleStepType.OUTCOME);
    expect(firstCase?.steps[1]?.source.exampleLocation?.line).toBeGreaterThan(0);
    expect(firstCase?.steps[0]?.source.sha256).toBe(feature.sourceSha256);
  });

  test("rejects tag contract violations", async () => {
    expect(() =>
      parseFeature(
        "features/opc/missing-id.feature",
        [
          "@implemented @bun",
          "Feature: invalid ids",
          "  Scenario: missing id",
          "    Given a package exists",
          "    Then the package contains /a.xml",
          "",
        ].join("\n"),
      ),
    ).toThrow(/exactly one @id-\*/i);

    const root = await makeProject({
      "features/opc/one.feature": [
        "@implemented @bun",
        "Feature: one",
        "  @id-dup",
        "  Scenario: first",
        "    Given a package exists",
        "    Then the package contains /a.xml",
        "",
      ].join("\n"),
      "features/opc/two.feature": [
        "@planned",
        "Feature: two",
        "  @id-dup",
        "  Scenario: second",
        "    Given a package exists",
        "    Then the package contains /b.xml",
        "",
      ].join("\n"),
    });

    let error: unknown;
    try {
      await inventoryFeatures(root);
    } catch (caught) {
      error = caught;
    }
    expect(String(error)).toContain("Duplicate scenario id @id-dup");
  });

  test("fails undefined bindings without silent pass", async () => {
    const root = await makeProject({
      "features/opc/undefined.feature": [
        "@implemented @bun",
        "Feature: undefined binding",
        "  @id-undefined",
        "  Scenario: unbound step",
        "    Given a package exists",
        "    Then the package contains /a.xml",
        "",
      ].join("\n"),
    });

    let error: unknown;
    try {
      await runAcceptance([], { root });
    } catch (caught) {
      error = caught;
    }

    expect(String(error)).toContain("Undefined step: a package exists");
    const artifact = (await Bun.file(join(root, "artifacts", "acceptance.json")).json()) as {
      status: string;
      execution: { steps: { undefined: number; skipped: number } };
    };
    expect(artifact.status).toBe("failed");
    expect(artifact.execution.steps.undefined).toBe(1);
    expect(artifact.execution.steps.skipped).toBe(1);
  });

  test("fails ambiguous bindings", async () => {
    const root = await makeProject({
      "features/opc/ambiguous.feature": [
        "@implemented @bun",
        "Feature: ambiguous binding",
        "  @id-ambiguous",
        "  Scenario: duplicate match",
        "    Then I see 1 part",
        "",
      ].join("\n"),
    });

    const bindings: StepBinding[] = [
      { pattern: /I see (\d+) part/, run: () => undefined },
      { pattern: /I see (\d+) part/, run: () => undefined },
    ];

    let error: unknown;
    try {
      await runAcceptance(bindings, { root });
    } catch (caught) {
      error = caught;
    }

    expect(String(error)).toContain("Ambiguous step: I see 1 part");
    const artifact = (await Bun.file(join(root, "artifacts", "acceptance.json")).json()) as {
      execution: { steps: { ambiguous: number } };
    };
    expect(artifact.execution.steps.ambiguous).toBe(1);
  });

  test("rejects scenarios without Then and reports failing Then assertions", async () => {
    expect(() =>
      parseFeature(
        "features/opc/no-then.feature",
        [
          "@implemented @bun",
          "Feature: empty pass guard",
          "  @id-no-then",
          "  Scenario: only setup",
          "    Given a package exists",
          "    When I append /a.xml",
          "",
        ].join("\n"),
      ),
    ).toThrow(/at least one Then step/i);

    const root = await makeProject({
      "features/opc/failing-then.feature": [
        "@implemented @bun",
        "Feature: failing assertions",
        "  @id-failing-then",
        "  Scenario: failed outcome",
        "    Given a package exists",
        "    Then the package contains /missing.xml",
        "",
      ].join("\n"),
    });

    const bindings: StepBinding[] = [
      { pattern: /a package exists/, run: (ctx) => void (ctx.state = { ready: true }) },
      {
        pattern: /the package contains (.+)/,
        run: (_ctx, part) => {
          throw new Error(`missing ${part}`);
        },
      },
    ];

    let error: unknown;
    try {
      await runAcceptance(bindings, { root });
    } catch (caught) {
      error = caught;
    }

    expect(String(error)).toContain("missing /missing.xml");
    const artifact = (await Bun.file(join(root, "artifacts", "acceptance.json")).json()) as {
      execution: { cases: { failed: number }; steps: { failed: number } };
    };
    expect(artifact.execution.cases.failed).toBe(1);
    expect(artifact.execution.steps.failed).toBe(1);
  });

  test("rewrites stale acceptance artifacts before a failing run", async () => {
    const root = await makeProject({
      "features/opc/stale.feature": [
        "@implemented @bun",
        "Feature: stale artifact reset",
        "  @id-stale",
        "  Scenario: stale report",
        "    Then I fail now",
        "",
      ].join("\n"),
      "artifacts/acceptance.json": JSON.stringify({
        runId: "stale-run",
        status: "passed",
        failures: [],
      }),
    });

    const bindings: StepBinding[] = [
      {
        pattern: /I fail now/,
        run: () => {
          throw new Error("boom");
        },
      },
    ];

    let error: unknown;
    try {
      await runAcceptance(bindings, { root });
    } catch (caught) {
      error = caught;
    }

    expect(String(error)).toContain("boom");
    const artifact = (await Bun.file(join(root, "artifacts", "acceptance.json")).json()) as {
      runId: string;
      status: string;
      failures: string[];
    };
    expect(artifact.runId).not.toBe("stale-run");
    expect(artifact.status).toBe("failed");
    expect(artifact.failures[0]).toContain("boom");
  });

  test("does not require a steps module when nothing executes", async () => {
    const root = await makeProject({
      "features/planned/deferred.feature": [
        "@planned",
        "Feature: deferred work",
        "  @id-planned",
        "  Scenario: inventory only",
        "    Given a package exists",
        "    Then the package contains /later.xml",
        "",
      ].join("\n"),
    });

    const report = await runAcceptance(undefined, { root });
    expect(report.status).toBe("passed");
    expect(report.inventory.features.planned).toBe(1);
    expect(report.execution.features.planned).toBe(1);
  });
});

async function makeProject(files: Record<string, string>): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-gherkin-"));
  roots.push(root);
  await mkdir(join(root, "features"), { recursive: true });
  for (const [relativePath, content] of Object.entries(files)) {
    const path = join(root, relativePath);
    await mkdir(dirname(path), { recursive: true });
    await Bun.write(path, content);
  }
  return root;
}
