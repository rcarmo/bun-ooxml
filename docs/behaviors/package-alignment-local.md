## Package alignment local checks

Bun executes the sealed second batch through its production ZIP, OPC, graph, comparison, Presentation and Workbook APIs: 20 IDs, 27 cases and 127 steps. The candidate has local results only; the published reference stays at v0.152.0 and `28e492f50979aaec6ab8d8d001cd9c37790e7fc6`.

[`package-alignment-candidate.json`](package-alignment-candidate.json) records the exact root, shared commit, manifest hash, 62 feature seals and four Office outline identity migrations. Reference verification rejects dirty files, hidden tracked-byte changes, wrong root/commit/manifest and feature-seal drift. Historical identity checks reverse only those reviewed outline changes. Earlier XML and runtime candidates keep their separate metadata.

The production change is in `src/xml/comparison.ts`. OPC `Relationship` attributes `Type` and `Target` contain URI strings. Their literal values participate in comparison without QName prefix lookup; unknown QName-like attributes elsewhere retain conservative namespace checks. Changed Type/Target values, duplicate IDs, external-mode changes and QName rebinding controls stay different.

`tests/acceptance/package-alignment.ts` adds the concrete recipes without replacing historical bindings. It verifies all eleven ZIP refusals, five independent pre-expansion budgets against valid and corrupt-DEFLATE siblings, mixed Store/Deflate writing, transactional custody, exact MIME-only diff sets and native Office edits. Corpus no-op saves run against all 36 Go and 35 Python fixtures. XLSX readback includes the requested string and effective wrap-text style, and both Office editors preserve the original `link:id` values.

Checks run from the repository root:

```sh
make check
OOXML_FIXTURES_ROOT=/workspace/projects/fixtures-ooxml \
OOXML_REFERENCE_PIN=/workspace/projects/bun-ooxml/docs/behaviors/package-alignment-candidate.json \
make check
```

The default lane passes 1,124 unit tests with 10 candidate-only skips. The candidate lane passes 1,134 unit tests with no skips. Both pass 732/732 implemented acceptance cases, all three save/reopen examples, reference verification and shared input checks.

Six temporary production faults fail the exact batch and are restored before full checks: CRC vector, archive budget, relationship-URI interpretation, transaction rollback, MIME-only diff and writer UTF-8 flag. Logs and exact per-step reports are under `/workspace/analysis/ooxml-alignment-next20-20260930/`.

ZIP64 has no new contract or implementation in this batch. Go and Python have independent preparation and review work; their execution cannot be inferred from Bun's results. No push, tag, release, pin advancement or shared execution credit is authorised.
