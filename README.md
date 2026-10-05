# bun-ooxml

Create and edit DOCX, PPTX and XLSX files from TypeScript running on Bun. The package writes Office Open XML archives in-process, without launching Office or LibreOffice, and aims to leave package parts outside an edit untouched. It is useful when a document change needs an exact target, a saved file and a check that unrelated content survived.

This is a bounded editor, not a general-purpose Office replacement. Word Compare, slide imports, chart authoring, merged-table restructuring, structural spreadsheet edits and formula calculation are not implemented. [Supported operations and limits][scope] has the current detail.

## Try it

Use Bun 1.4.1 or newer. Clone with the shared fixtures submodule, then install and run the examples:

```bash
git clone --recurse-submodules https://github.com/rcarmo/bun-ooxml.git
cd bun-ooxml
make install
make examples
```

The example creates a document, a presentation and a workbook, reopens each file and checks its content. For editing existing files, see [`examples/agent-edit.ts`](examples/agent-edit.ts) and the [usage guide][usage]. If you cloned without `--recurse-submodules`, run `git submodule update --init --recursive`; missing fixture inputs fail instead of being fabricated.

A guarded Word text edit can be as small as this:

```ts
import { Document } from "bun-ooxml";

const document = await Document.open("contract.docx");
const matches = document.find("thirty calendar days");
if (matches.length !== 1) throw new Error("Expected exactly one target");
matches[0]!.replace("thirty business days");
await document.save("contract-edited.docx");
```

Save to a new path, reopen the result and inspect the affected parts before using it elsewhere. The [workflow API][workflow] adds previews, strict batches, guarded saves and per-target receipts when a single edit is too simple for the job.

## What you can edit

Word supports document creation, paragraphs, runs and tables; existing-file edits include [text and run formatting][run], [paragraph and page properties][page], [style authoring][styles], [core metadata][metadata] and selected table properties. Linked-story review, bounded tracked replacements and existing comment-thread inspection cover specific review tasks. They do not implement general Word Compare or comment authoring.

PowerPoint supports presentations with title slides, tables and [positioned text boxes][textbox]. Existing decks support [slide reordering][slides], slide text and [speaker-note edits][notes]. [Effective bold and italic inspection][formatting] follows a bounded paragraph-style cascade and reports unsupported contexts.

Excel supports sheets and cells, existing-file cell edits and [cell-style selection][cellstyle] that keeps values, formulas and caches. [A1 range parsing][ranges], [formula reference analysis][formula] and [string-level reference remapping][remap] operate on references; none calculates a workbook.

At the package level, [`admitPackage()`][admission] checks bounded ZIP structures and XML members. Guarded OPC graph edits add or detach parts and relationships; [package diffs][diff] separate equivalent XML from changed binary members, additions and removals. The [XML snapshot APIs][xml] make bounded, source-preserving changes. ZIP, XML and OPC runtime code uses Bun's file, hash and compression APIs; there are no runtime package dependencies or Office subprocesses.

## Contracts and checks

Development commands use Make targets (or `bun scripts/dev-run.ts <label>
run|test|tool <arguments>`). Host caches/build/scratch live under
`/workspace/tmp/bun-ooxml/{cache/<tool>,build,runs/<purpose>/<run-id>}` when
writable locally, otherwise the platform temp directory. Explicit absolute
`PROJECT_TMP_BASE` selects its `bun-ooxml` child; compatible `PROJECT_TMP_ROOT`
must agree when both are set. CI tries `$RUNNER_TEMP`, original `$TMPDIR`, then
platform temp, always with a `bun-ooxml` child. Invalid overrides fail.
Stable `tests/` and `logs/` roots share this hierarchy.
Profiles and retained
reports stay in `artifacts/policy-profiles/`, separate from disposable paths.
`make clean CONFIRM_IDLE=yes` requires idle jobs and never removes artifacts.

The [`fixtures-ooxml` repository][fixtures] supplies shared workflow contracts, format facts and read-only fixtures. This repository pins it as a Git submodule and resolves fixture inputs by stable manifest ID. A shared scenario describes an expected outcome; it does not by itself prove that the Bun implementation executes it. The [test catalogue][tests] lists the native bindings and their current status.

```bash
make check
```

This runs type checks, inventory and mapping validation, unit tests, Gherkin acceptance tests and the examples. `make parity` also demands implementation of every specified behaviour, so it currently fails on unsupported operations. The separate `make office-oracles` target checks three authored files with Open XML SDK, LibreOffice and Poppler; those tools are development-only. `make property-campaign` runs fixed-seed XML/ZIP/Office properties and bounded timing workloads. See [independent checks][oracles] and [campaign scope][campaign] for prerequisites and limits.

MIT licensed. [Third-party notices](THIRD_PARTY_NOTICES.md) list licences and fixture provenance.

[scope]: docs/contracts/port-scope.md
[usage]: docs/agents/usage.md
[workflow]: docs/contracts/workflow-api.md
[run]: docs/contracts/run-formatting.md
[page]: docs/contracts/page-layout.md
[styles]: docs/contracts/style-authoring.md
[metadata]: docs/contracts/core-properties.md
[textbox]: docs/contracts/text-box.md
[slides]: docs/contracts/slide-order.md
[notes]: docs/contracts/notes-editing.md
[formatting]: docs/contracts/effective-formatting.md
[cellstyle]: docs/contracts/cell-style.md
[ranges]: docs/contracts/a1-ranges.md
[formula]: docs/contracts/formula-analysis.md
[remap]: docs/contracts/formula-remap.md
[admission]: docs/contracts/package-admission.md
[diff]: docs/contracts/package-comparison.md
[xml]: docs/contracts/xml-byte-snapshot.md
[fixtures]: https://github.com/rcarmo/fixtures-ooxml
[tests]: docs/behaviors/README.md
[oracles]: docs/contracts/office-oracles.md
[campaign]: docs/contracts/property-campaign.md
