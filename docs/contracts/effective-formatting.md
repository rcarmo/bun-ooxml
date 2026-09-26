# Effective bold and italic

`Paragraph.effectiveRunFormatting()` returns one entry per direct run, with its
text, zero-based run index, paragraph-style chain and effective bold/italic flags.
Each flag includes ordered contributions with source, part, optional style ID,
operation, supplied value and resulting state.

This inspection supports plain direct body paragraphs containing Latin-range
text and punctuation. It reads current style definitions without editing bytes,
invalidating handles or caching results. Returned arrays and objects are detached.
An out-of-band main-document change makes the paragraph handle stale.

The cascade starts at false, applies explicit `docDefaults/rPrDefault` flags,
then the selected paragraph style's base chain from root to leaf, and finally
absolute direct run flags. Without a direct paragraph style, the default paragraph
style is selected. An explicit style does not also inherit the default style
unless its `basedOn` chain names it.

Bold and italic are OOXML toggle properties: a true value in a style reverses the
state inherited so far; false leaves it unchanged. Direct run values set the
result absolutely. Provenance uses `set`, `toggle` and `retain` operations.
See [the bold definition](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_b_topic_ID0EP6EO.html)
and [Open XML SDK's bold reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.bold).

The inspector refuses table/nested contexts, character or linked styles,
numbering, nonempty paragraph properties other than direct `pStyle`, fields,
revisions, drawings, complex-script flags and unknown run properties. Text outside
the bounded Latin/combining-mark/punctuation ranges also refuses. Font, theme,
script, language, table and numbering cascades are not implemented. A paragraph
mark's run properties are outside this direct-text-run inspection.

Styles must have an internal, uniquely resolved relationship and correct MIME.
IDs must be unique across the registry, and the paragraph default must be unique.
Only the selected ancestry is traversed; missing, cyclic or nonparagraph bases
refuse. Unused ancestry is not evaluated. Office `stylesWithEffects` relationships
refuse because the second registry is not reconciled. Read-only inspection is
allowed on protected documents.

## Independent comparison

Open XML SDK validates the authored probe. LibreOffice 24.2.7 PDF output, extracted
with Poppler, agrees on inherited bold, explicit direct-off, and direct bold plus
italic. It retains bold across two true paragraph-style toggles, while the OOXML
rule above returns false. The report records **3/4 matching markers** and the
known difference. The optional oracle fails on unexpected mismatches in the other
markers; it does not treat this known difference as agreement. Microsoft Word and
general font/script/layout equivalence have not been validated.

Canonical contract: `workflows/docx/effective-formatting.feature` in the shared
reference. Native checks: `tests/unit/docx-effective-formatting.test.ts` and
`tests/acceptance/effective-formatting.ts`. Their coverage includes provenance,
read-only custody, retained spans, namespace/UTF-16 handling, live style changes,
ambiguous registries, unsupported contexts and saved/reopened values.
