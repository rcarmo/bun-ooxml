# Independent authored-file checks

`make office-oracles` generates three small native Office files and checks them
with independent tools. It is a development-only lane; production operations and
`make check` remain Bun-only. CI runs the lane separately and uploads its inputs,
PDFs, extracted text, recalculated workbook and JSON report.

Install .NET 10, LibreOffice Writer/Impress/Calc and Poppler. The validator pins
Open XML SDK 3.5.1 and transitive packages through `packages.lock.json`; restore
uses locked mode. LibreOffice and Poppler versions are recorded on every run.
Ubuntu CI packages receive distribution updates, so rendering binaries are not
reproducibly pinned. PDF hashes are provenance, not stable golden images.

```sh
sudo apt-get install --no-install-recommends libreoffice-writer libreoffice-impress libreoffice-calc poppler-utils
make office-oracles
```

## Exact scope

* Open XML SDK validates the generated DOCX, PPTX and XLSX against its Office2019
  profile. A deliberately broken PPTX must fail on missing view metadata, which
  prevents a non-running validator from appearing successful. Inputs are opened
  read-only and checked for unchanged hashes.
* LibreOffice exports each sample to an isolated PDF directory with a fresh
  temporary user profile. Poppler checks expected page counts (1, 2, 1), text
  markers and the two slide titles on their expected pages. These checks detected
  missing title text; they do not measure clipping in general, compare pixels,
  measure all bounds or establish font/layout equivalence with Microsoft Office.
* The XLSX sample starts with a deliberately stale `SUM(A1:A2)` cache. A native
  Bun edit changes A1 from 20 to 21 and removes the cached answer. LibreOffice must
  render and save the formula result 43 from 21+22. Bun then reads that result from
  the independently saved XLSX. Formula assembly uses native OPC in this test;
  it is not a new formula-authoring or calculation API.

The initial schema probe found three incomplete view-property elements in newly
created PPTX files. Native regression tests and the SDK check now cover pane
sizes, scales and origins. The initial PDF probe also found clipped titles caused
by absent placeholder geometry; the owned title layout now supplies nonoverlapping
title/subtitle bounds. Existing imported layouts are unchanged.

A separate initial probe imported stale XLSX cache 999 into LibreOffice 24.2.7 and retained 999 despite
recalculation flags. The successful lane depends on Bun's explicit cache removal.
Applications may choose different recalculation policies; flags alone do not prove
that cached answers are fresh.

## Unverified work

The lane covers three generated samples, not the full fixture corpus, arbitrary
layout inheritance, revisions/comments, pivots, charts or formula grammar.
Microsoft Word/PowerPoint/Excel rendering, pixel comparisons, broad calculation
parity, fuzzing and performance remain open. No schema or rendering result supplies
execution credit to sibling consumers or closes the full release gate.

Each oracle command runs in a separate POSIX process group, with a timeout,
a 4 MiB output cap and a bounded post-kill drain. The timeout test starts a real
descendant holding the pipes and checks that it cannot write a delayed marker.
Independent validation does not run on Windows in this lane.

Sources: `scripts/office-oracles.ts`, `scripts/oracle-process.ts`, `tests/oracles/schema/`, and the two native
regressions in `tests/unit/pptx-create.test.ts`. Results are written to
`artifacts/office-oracles/report.json`, including failed runs.
