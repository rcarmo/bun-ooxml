# Retained PowerPoint styling and Word formatting

The explicit shared candidate at `7f9a3469d335b01d30b6bbd8acdecc490bc9f0d0` adds 20 PPTX cases/201 steps and 20 Word cases/200 steps. The default fixture pin is unchanged. Shared declarations remain planned with `executionCredit:false`.

## APIs

PowerPoint slide handles expose `patchShapeStyle(shapeId, patch)`, `patchTextFrame(shapeId, patch)` and `patchRetainedRunEffects(shapeId, paragraph, run, patch)`. They select an existing ordinary ungrouped shape, splice its direct properties, validate the package transaction and retain text, geometry, notes and relationships. Explicit noFill differs from removing a direct fill.

Word paragraph handles expose `patchRetainedRun(run, patch)` and `patchRetainedProperties(patch)`. The run method changes one selected direct text run; legacy `formatRuns` keeps its all-runs semantics. Word edits retain all other run/paragraph/member bytes, refuse tracking/protection or unsupported content, and invalidate old paragraph handles only after a successful commit. Font size uses integer half-points; paired ASCII/high-ANSI font edits leave unrelated script slots outside the bounded profile. Line480/auto means double line spacing in 240ths of a line.

## Verification

`tests/acceptance/retained-style-word.ts` binds all40 sealed IDs with exact patch operands and independently parses saved ZIP/XML. Assertions check property presence/values, direct schema order, whole unpatched XML bytes, original member sets/payloads, caller input and unchanged relationships/content types. The custody scanner consumes complete quoted attribute tokens. Its property-mask oracle is independent of production splice code.

`tests/unit/retained-style-word-native.test.ts` covers invalid multifield patches, types/bounds, executable patches, missing/self-closing properties, quoted tokens, theme/alpha/duplicate/protected/tracked/field/hyperlink profiles, stale handles, staged rollback and saved-output failure. Production family no-op and unpatched-language faults fail acceptance, then pass after exact source restoration. Native test counts and inventory mappings are not coverage percentages.

Use the reviewed `docs/behaviors/retained-style-word-candidate.json` with explicit `OOXML_FIXTURES_ROOT` and `OOXML_REFERENCE_PIN`. This lane executes the earlier manipulation and formatting cases as well as the new40; default execution keeps its existing lane. Historical test helpers remove only exact additive IDs and sealed predicates before checking old receipts. No publication, central adoption, renderer, theme, chart, table, slide-deletion/import or default-pin claim follows from candidate execution.
