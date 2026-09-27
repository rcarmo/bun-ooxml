import { afterAll, describe, expect, test } from "bun:test";
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import { PickleStepType } from "@cucumber/messages";

import { executeAcceptance, newAcceptanceRunId, parseFeature, type AcceptanceFeature, type AcceptanceInventory } from "../../scripts/gherkin.ts";
import { bindings, cleanupWorkflowFixtures } from "../acceptance/workflow.ts";

const PROJECT_ROOT = new URL("../..", import.meta.url).pathname;
import {sharedScenarios} from '../helpers/shared-scenarios.ts';

afterAll(async () => {
  await cleanupWorkflowFixtures();
});

describe("workflow shared acceptance", () => {
  test("executes the frozen shared workflow feature as 19 implemented Bun cases", async () => {
    const contract=await Bun.file(join(fixturesRoot(),'contracts/mutation-safety.json')).json();
    const features=await sharedScenarios(contract.scenarioIds),scenarios=features.flatMap(f=>f.scenarios);
    const inventory = inventoryFor(features);
    const caseCount = scenarios.flatMap((scenario) => scenario.cases).length;
    const expectedOutcomeCount = scenarios
      .flatMap((scenario) => scenario.cases)
      .flatMap((acceptanceCase) => acceptanceCase.steps)
      .filter((step) => step.type === PickleStepType.OUTCOME).length;

    expect(features.every(f=>f.lifecycle==='implemented')).toBe(true);
    expect(features.every(f=>f.runner==='bun')).toBe(true);
    expect(features.map(f=>f.path).sort()).toEqual([...contract.features].sort());
    expect(features.every(f=>!f.path.includes('/planned/'))).toBe(true);
    expect(scenarios).toHaveLength(8);
    expect(caseCount).toBe(19);

    const execution = await executeAcceptance(inventory, bindings, newAcceptanceRunId());
    const passedOutcomeCount = execution.features
      .flatMap((report) => report.scenarios)
      .flatMap((report) => report.cases)
      .flatMap((report) => report.steps)
      .filter((step) => step.type === PickleStepType.OUTCOME && step.status === "passed").length;

    expect(execution.failures).toEqual([]);
    expect(execution.counts.features).toEqual({ passed: 5, failed: 0, planned: 0, total: 5 });
    expect(execution.counts.scenarios).toEqual({ passed: 8, failed: 0, planned: 0, total: 8 });
    expect(execution.counts.cases.passed).toBe(19);
    expect(execution.counts.cases.failed).toBe(0);
    expect(execution.counts.cases.planned).toBe(0);
    expect(execution.counts.cases.total).toBe(19);
    expect(passedOutcomeCount).toBe(expectedOutcomeCount);
    expect(execution.counts.steps.failed).toBe(0);
    expect(execution.counts.steps.undefined).toBe(0);
    expect(execution.counts.steps.ambiguous).toBe(0);
  }, 300_000);
});

function inventoryFor(features: AcceptanceFeature[]): AcceptanceInventory {
  const selected=features.flatMap(f=>f.scenarios);
  const scenarios = selected.length;
  const cases = selected.flatMap((scenario) => scenario.cases).length;
  const steps = selected
    .flatMap((scenario) => scenario.cases)
    .flatMap((acceptanceCase) => acceptanceCase.steps).length;
  return {
    root: PROJECT_ROOT,
    features,
    counts: {
      features: { implemented: features.length, planned: 0, total: features.length },
      scenarios: { implemented: scenarios, planned: 0, total: scenarios },
      cases: { implemented: cases, planned: 0, total: cases },
      steps: { implemented: steps, planned: 0, total: steps },
    },
  };
}
