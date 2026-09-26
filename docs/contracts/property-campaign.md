# Seeded properties and bounded timings

`make property-campaign` runs two deterministic synthetic corpora and five fixed
native timing workloads. Production code and shared fixture bytes are untouched.
CI runs these checks separately and uploads their corpora and JSON reports.

Each seed (`20260926`, `8675309`) generates 256 inputs with a 32-bit LCG. Exact
UTF-8 JSON bytes are written to `artifacts/property-campaign/seed-<seed>/corpus.json`
and hashed in its report. The CLI accepts a seed, 1–2,000 cases and 1–100 Office
cases, with Office cases no greater than the corpus count. The default Make target
uses 32 Office cases per format and seed.

## Tested properties

* XML uses synthetic prefixed leaves with escaped text/attributes, Unicode,
  whitespace and metacharacters. Parsing and successful edits must recover exact
  values; accepted edits reparse and retain untouched lexical text. Each source
  also receives a negative bound, an overlap and a forbidden DTD insertion; each
  must return the specific typed error. Texts contain at most 59 generated symbols.
* ZIP32 and forced-ZIP64 archives each contain three members with payloads up to
  255 bytes. Original payloads and deterministic encoding roundtrip. Four seeded
  single-bit flips per archive exercise admission. A typed `zip-*` error counts
  as refusal; untyped exceptions fail. Accepted mutations must survive a further
  write/read roundtrip. Both accepted and refused reads must leave the mutated
  caller buffer unchanged. A tolerated metadata change is not automatically a bug.
  Reader limits cap entries, per-entry/total output, input size and expansion ratio.
* DOCX text replacement/formatting, PPTX text-box append and XLSX direct-style
  selection/removal are saved and reopened. Text, selected values/styles and
  unrelated package parts are checked; stale DOCX spans and PPTX anchors must
  refuse without changing the archive. XLSX invalid-style refusal also preserves
  bytes. These samples do not validate every property or imported topology.

Failures carry the seed, case identity, exact input and error. ZIP mutation failures
also record bit position and archive/mutant hashes. The full corpus makes a run
replayable. `report.json` records completed-case counters, errors, refusal codes,
source hashes and environment. Status and the failure list determine success;
completed-operation counters alone do not, including failures injected after a
completed operation. Reports get a fresh run ID and `running` state before work;
failed runs replace earlier successful reports. Counts describe only the sampled
properties.

## Timing method

The performance script creates fixed inputs before timing. Each workload gets one
warm-up and seven measured runs. Reports keep every sample, median and nearest-rank
p95; with seven samples, p95 is the maximum. Timings include the named native
operations and correctness checks. They exclude input construction and disk I/O.

Workloads: parse/edit/reparse 1,001 XML elements; read/write/read a 128-member ZIP64
archive with 512 KiB of uncompressed payload; open/edit/save/reopen a 50-paragraph
DOCX, an eight-slide PPTX and a 128-cell XLSX. Each Office sample changes one target.
Process RSS is sampled after all workloads; it is not peak memory or allocation cost.
GC and CPU scheduling are uncontrolled. There is no speedup comparison or portable
pass threshold.

These are fixed-seed property samples, not a coverage-guided fuzzing campaign.
No minimiser, all-input proof, large-archive stress, schema validation or independent
Office execution is included. See [independent validation](office-oracles.md) for
the separate three-sample schema and PDF checks.
