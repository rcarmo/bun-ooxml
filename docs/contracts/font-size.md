# Direct Word font size

`paragraph.setRunFormatting({ fontSizePt: 10.5 })` writes a direct `w:sz` value
of `21` on each supported run in that paragraph. `paragraph.directFontSizes()`
returns those direct sizes in run order, using `null` where no direct `w:sz`
exists. It does not compute inherited size or rendered appearance.

```ts
const document = await Document.open('input.docx');
const paragraph = document.paragraphs[0]!;
paragraph.setRunFormatting({ fontSizePt: 10.5 });
await document.save('output.docx');
const reopened = await Document.open('output.docx');
console.log(reopened.paragraphs[0]!.directFontSizes());
```

ECMA-376 Part 1 (2016), §17.3.2.38 (`sz`, PDF page 323, printed page 313), defines
non-complex-script size in half-points. §17.18.42 (`ST_HpsMeasure`, PDF pages
1406–1407, printed pages 1396–1397) also permits unit-bearing positive measures.
The numeric schema branch is `xsd:unsignedLong`; its lexical zero is rejected by
this editor's positive-size policy. This editor accepts only
positive point values whose doubled value is a JavaScript safe integer. It
neither rounds quarter-points nor converts unit-bearing stored values. This is
an API limit, not the complete schema range.

`undefined` leaves the property alone; `null` removes the direct override.
Removing an override can expose an inherited value, but the direct getter still
returns `null`. Equal sizes are lexical no-ops, including a stored leading-zero
half-point spelling. Size can be combined with bold and italic in one patch.
Existing `w:szCs` bytes are preserved; complex-script sizing and style evaluation
are separate operations.

The [run-formatting rules](run-formatting.md) apply: supported direct text runs
only, preflight every run before mutation, preserve unrelated bytes and encoding,
and refuse protection, stale handles or ambiguous run properties. Real changes
invalidate paragraph/span snapshots. Validation or serialization failure rolls
back the attempted change and leaves prior valid edits and handles intact.

The shared half-point scenario checks actual save/reopen, unchanged text, exactly
one direct `w:sz w:val="21"`, and a reopened direct size of 10.5 points. It does
not check glyph rendering, font substitution or whole-document style resolution.
