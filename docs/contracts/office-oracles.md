# Independent authored-file checks

`make office-oracles` generates three small native Office files and checks them
with independent tools. These development checks are separate from the Bun-only
production APIs and `make check`. CI uploads their inputs,
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
  markers, the two slide titles on their expected pages after reversing slide
  order, and the DOCX sample's
  792×612-point landscape page dimensions. They do not measure clipping in general, compare pixels,
  measure all bounds or establish font/layout equivalence with Microsoft Office.
* Four Latin DOCX markers compare native bold/italic with Poppler PDF markup.
  LibreOffice 24.2.7 agrees on three; it keeps bold after two true paragraph-style
  toggles where OOXML evaluation returns false. This known difference is recorded
  separately; unexpected marker differences fail. Microsoft Word and general
  font/script agreement remain unverified.
* The XLSX sample starts with a deliberately stale `SUM(A1:A2)` cache. A native
  Bun edit changes A1 from 20 to 21 and removes the cached answer. LibreOffice must
  render and save the formula result 43 from 21+22. Bun then reads that result from
  the independently saved XLSX. Formula assembly uses native OPC in this test;
  it is not a new formula-authoring or calculation API.

The SDK check covers PPTX view-property pane sizes, scales and origins. Generated
title layouts have explicit nonoverlapping title/subtitle bounds; imported layouts
are unchanged.

LibreOffice 24.2.7 retained a stale XLSX cache value of 999 when imported with
recalculation flags alone. The successful recalculation check requires Bun's
explicit cache removal.
Applications may choose different recalculation policies; flags alone do not prove
that cached answers are fresh.

## Limits

These checks cover three generated samples. They do not cover the full fixture
corpus, arbitrary layout inheritance, revisions/comments, pivots, charts or formula
grammar. Microsoft Word/PowerPoint/Excel rendering, pixel comparison and broad
calculation accuracy are unverified.

Each oracle command runs in a separate POSIX process group, with a timeout,
a 4 MiB output cap and a bounded post-kill drain. The timeout test starts a real
descendant holding the pipes and checks that it cannot write a delayed marker.
These independent checks do not run on Windows.

Sources: `scripts/office-oracles.ts`, `scripts/oracle-process.ts`, `tests/oracles/schema/`, and the two native
regressions in `tests/unit/pptx-create.test.ts`. Results are written to
`artifacts/office-oracles/report.json`, including failed runs.
