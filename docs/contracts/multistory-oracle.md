# Independent multistory revision evidence

`make office-oracles` compares native revision resolution with separately authored
Open XML SDK documents covering body, header, footer and footnotes. This optional
development gate does not run in library operations or ordinary acceptance.

## Producer and reader

`tests/oracles/schema/MultistoryRevisions.cs` uses DocumentFormat.OpenXml 3.5.1
strongly typed classes to author four documents:

- a source containing insertion, deletion, direct run-property change and paired
  run moves in each of the four stories;
- accepted and rejected expectations, built directly from final paragraphs and
  properties without resolving the source;
- a schema-valid negative source with an unsupported paragraph-property change
  in the last selected footnote story.

The source has five revision records per story, or twenty in total. A move has
two wrapper records. The expected documents have no revisions. The negative
source has an additional `pPrChange`, for twenty-one SDK revision records.
Footnote references, separator and continuation separator parts are present.
Relationships are authored by the SDK. Neither the producer nor its reader calls
Bun's parser, editing APIs or synthetic-fixture helpers.

The same SDK validates all four documents and native outputs against its Office2019
schema profile. Its reader reports ordered paragraphs, runs, text, bold/italic
flags and revision identities for each story. `scripts/multistory-oracle.ts`
compares this report against literal expectations defined separately in TypeScript.
The reader protocol requires the expected SDK version, filename, SHA-256, complete
story inventory, successful schema status and no truncation or exceptions.

## Native comparison

The Bun resolver uses `text-properties-and-moves` and explicitly selects the four
named story parts. Both actions must return twenty resolved records and the exact
changed-part list. Saved output story bytes must equal the separately authored
expected document's story bytes. Every other member must equal the original
source. All source and expected archive files retain their SHA-256 values.

Reopening the native output must show no supported or unsupported revisions.
Repeating the operation must report zero changes and preserve exact current
archive bytes. With the unsupported footnote source, both actions must report
`docx-revisions-unsupported` and leave the complete package unchanged, including
all earlier selected story content.

Three schema-valid corruptions change header text, formatting or move order.
The SDK reader must still validate each file's schema but the literal semantic
comparison must reject it. Unit controls separately reject wrong reader/file/hash
identity, failed or truncated schema evidence, missing stories, wrong content,
formatting, order and leftover revisions.

## Scope

The output is recorded under `artifacts/office-oracles/multistory` and in the main
oracle report with producer/reader versions, source-code hashes and per-file hashes.
The SDK authors and reads OOXML; it is not Microsoft Word. This gate does not
compare rendering or exercise arbitrary revision types or independent interactive
accept/reject behaviour.

The original `@id-docx-review-multistory-resolution` scenario remains planned.
This evidence needs independent review, central fixture/provenance registration
and explicit outcome bindings before its lifecycle can change. Existing shared
execution and mappings are unaffected.
