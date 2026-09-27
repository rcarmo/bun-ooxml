# Direct Latin run fonts

`Paragraph.setRunFormatting({ fontName })` sets matching `w:ascii` and `w:hAnsi`
attributes on each supported direct run's `w:rFonts` element.
`Paragraph.directFontNames()` returns those names in run order, with null for an
absent `rFonts` element.

```ts
const paragraph = document.paragraphs[0]!;
paragraph.setRunFormatting({ fontName: 'Times New Roman', fontSizePt: 12 });
const names = document.paragraphs[0]!.directFontNames();
// ['Times New Roman'] for a paragraph with one run
```

Names must be nonempty, trimmed strings with at most 255 UTF-16 code units, without
C0/C1 control characters (U+0000–001F and U+007F–009F) or invalid XML characters. This is a bounded editor policy,
not a font-installation check. Unicode and XML metacharacters are accepted and
escaped without changing the decoded name.

Null removes the complete supported `rFonts` element. Undefined leaves it alone.
An identical decoded name preserves the element's original prefix, attribute
quotes and whitespace, as well as exact archive bytes. A real change expires
paragraph handles; reacquire the paragraph afterwards. The returned array is a
detached value and cannot change the package.

## Refusals and preservation

Reading or selecting an existing font element requires both direct Latin slots
to be present, valid and equal. Extra font attributes refuse, including theme
references, East Asian and complex-script slots, and font-selection hints.
Incomplete, unequal or wrong-namespace slots also refuse. These checks apply to
same-name assignments and removal requests, so a write cannot silently discard
unsupported metadata. Unrelated edits, such as bold, leave such font elements
untouched under the existing formatting guards.

The existing plain-run, protection, stale-target and transaction checks apply.
All selected runs pass preflight before mutation. A late-run refusal, staged
write failure or serialization failure restores package bytes and model handles.
Tests cover UTF-16 path save/reopen, table-grid preservation and unchanged
unrelated member payloads. This API does not add a font table or embed fonts.

## Specification and limits

ECMA-376 Part 1, fifth edition (October 2016), §17.3.2.26 defines `rFonts` and its
four font slots. Font selection depends on script classification and other
properties; theme and inherited settings can take precedence. The complete PDF
is in the [shared specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).

This getter reports matching direct ASCII/high-ANSI slots only. It does not
resolve theme fonts, inheritance, script selection, substitutions or availability.
It does not change the East Asian or complex-script font settings.

The six shared getter cases cover Arial, Times New Roman, Calibri, Courier New,
Georgia and Verdana in memory. Native tests separately cover saved values and
custody. One authored sample passes Open XML SDK's Office2019 schema profile;
visual appearance and installed-font behaviour have not been tested.
