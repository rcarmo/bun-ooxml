# Acceptance contract

The Bun-native acceptance runner lives in:

- `scripts/gherkin.ts`
- `scripts/acceptance.ts`

It uses the official `@cucumber/gherkin` parser/compiler plus
`@cucumber/messages` pickles. Acceptance is inventory-first: every compiled case
is known before execution, including Background expansion and Scenario Outline
rows.

## Source contract

Local regression features are loaded from `features/**/*.feature`.
`features/shared.json` additionally selects central features from the pinned
`references/fixtures-ooxml` checkout. Their lifecycle overlay is applied in memory;
no shared feature copy is written. Duplicate IDs across local and shared features
are rejected by the combined inventory.

The shared catalogue defines common expected behaviour. Native bindings and
per-consumer execution reports are local. Candidate mappings in `docs/behaviors/`
are not part of the execution inventory until reviewed, reconciled and bound.

Feature tags are strict:

- implemented features must be tagged exactly `@implemented @bun`
- planned features must be tagged exactly `@planned`

Scenario tags are strict:

- every scenario or scenario outline must declare exactly one `@id-*` tag
- scenario/example tags must not override lifecycle or runner tags
- example rows do not define IDs; expanded cases inherit the scenario ID and add
  exact example values to form a case identity

The parser rejects:

- empty feature names
- empty scenario names
- scenarios with no steps
- scenarios with no compiled cases
- scenarios without at least one `Then`/Outcome step
- duplicate feature names
- duplicate scenario names inside a feature
- duplicate scenario `@id-*` tags across the inventory
- duplicate expanded case identities
- malformed example headers

## Binding contract

`tests/acceptance/steps.ts` is loaded only when at least one implemented case
must execute.

It must export `bindings` (or a default array) with this shape:

```ts
export type StepBinding = {
  pattern: RegExp;
  run: (
    context: Record<string, unknown>,
    ...captures: string[]
  ) => unknown | Promise<unknown>;
};
```

Binding matches are exact against compiled step text. Partial regex matches do
not count.

Execution rules:

- no match => `undefined`
- more than one exact match => `ambiguous`
- thrown/rejected binding => `failed`
- later steps after a failure are marked `skipped`
- planned features are never executed; their steps are marked `planned`

## Runtime API

`runAcceptance()` is exported from `scripts/acceptance.ts`:

```ts
async function runAcceptance(
  bindings?: readonly StepBinding[],
  options?: {
    root?: string;
    full?: boolean;
    artifactPath?: string;
    stepsModulePath?: string;
  },
): Promise<AcceptanceRunReport>
```

If `bindings` is omitted, the runner imports `tests/acceptance/steps.ts` on
demand.

CLI:

```sh
bun run scripts/acceptance.ts
bun run scripts/acceptance.ts --full
```

`--full` fails whenever any planned feature remains.

## Artifact ledger

Every run rewrites `artifacts/acceptance.json` immediately with a fresh `runId`
so stale passing output cannot survive a later failure.

The final artifact records:

- unique `runId`
- start/finish timestamps
- root, artifact path and step-module path
- inventory counts split into `implemented` and `planned`
- execution counts for features/scenarios/cases/steps
- per-feature source SHA-256
- per-step source location and example-row location when applicable
- exact expanded step text and execution outcome
- all failures

The report is a scenario ledger, not just a pass/fail summary. Each expanded
Scenario Outline row is preserved as its own case with its exact values.
