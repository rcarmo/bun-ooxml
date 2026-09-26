# Native creation slice

All packages are built from TypeScript XML/OPC primitives, not copied fixtures or
external producer subprocesses. Generated files must reopen through native readers
and pass relationship/content-type checks. Visual/Office rendering remains unverified.

Parallel API contracts:

* DOCX `Document.create(): Document` and `addParagraph(text = '', options?: {bold?:boolean; italic?:boolean; style?:string}): Paragraph`.
  New package has main document and root relationship, body and section properties.
  Adds paragraph before sectPr; XML-escaped text and simple rPr bold/italic. Unknown
  style IDs refuse unless existing styles resolve (omit style support if unimplemented).
  Existing documents keep opaque parts, cache handles refreshed/stale on add.
* PPTX `Presentation.create(): Presentation` and `addTextSlide(title:string, subtitle?:string): Slide`.
  Author minimal valid presentation/master/layout/theme if needed; native title and
  subtitle placeholders, independent ID/relationship allocation. Existing decks
  without a compatible safe layout refuse rather than corrupt inheritance. Root
  fields slideId, dimensions etc valid. Document exact limitations.
* XLSX `Workbook.create(): Workbook`, `addWorksheet(name:string): Worksheet`,
  `Worksheet.setCellValue` extends to missing cells on created/existing supported
  plain sheets only with sorted rows/cells and dimension maintenance. Existing
  formulas/cache guards continue to apply; shared strings/styles/opaque content
  stay unchanged. New default Sheet1. Names: uniqueness case-insensitive, max31,
  forbidden characters and empty/apostrophe boundary names refuse. Writes bounded
  to Excel row/column maxima. Correct explicit inline strings, numeric/bool values.

Each format has src/*/index.ts ownership; supporting helper modules inside own
format directory allowed. Tests and features authoring first before code; bindings
exported tests/acceptance/create-{format}.ts, parent combines. No borrowed python
template binary, no font/render/runtime dependency. This is not all inherited
creation/style/theme/table/chart API parity.
