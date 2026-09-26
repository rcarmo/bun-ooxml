# Native package creation

The creation APIs build packages from TypeScript XML/OPC primitives. They do not
copy templates or invoke an external producer. Saved packages reopen through the
native readers and pass relationship/content-type validation. The
[independent checks](office-oracles.md) cover a few generated samples; general
Microsoft Office rendering compatibility remains unverified.

## Word

`Document.create()` returns an empty document. `addParagraph(text = '', options?)`
accepts `bold`, `italic` and an existing paragraph `style` ID. Text is escaped and
paragraphs are inserted before section properties. Unknown style IDs refuse.
Existing opaque parts are preserved; structural edits invalidate table/span
handles according to their documented lifetimes.

## PowerPoint

`Presentation.create()` creates an empty presentation with an owned
layout/master/theme graph. `addTextSlide(title, subtitle?)` adds title/subtitle
placeholders with independent slide and relationship IDs. Appending to an
existing deck requires a safe compatible layout; ambiguous or unsupported layout
inheritance refuses before mutation.

## Excel

`Workbook.create()` supplies `Sheet1` and default styles. `addWorksheet(name)`
adds a sheet while preserving existing relationships and opaque members.
`Worksheet.setCellValue()` can insert missing ordinary cells with sorted row/cell
order and updated dimension metadata. Formula-cache and unsupported-topology
checks still apply.

Sheet names are case-insensitively unique, no longer than 31 characters, and must
not be empty or contain forbidden characters or boundary apostrophes. Coordinates
stay within Excel row/column limits. Inline strings, numeric and boolean values
are written explicitly; formula assignment is outside this API.

`examples/create-office.ts` exercises all three save/reopen paths. Media, chart
and structural authoring beyond the documented APIs are unsupported.
