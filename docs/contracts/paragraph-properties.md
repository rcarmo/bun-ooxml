# Direct paragraph properties

`Paragraph.directProperties()` reads direct alignment, before/after spacing,
three paragraph flags and `outlineLevel`. `setProperties(patch)` changes those properties without
changing paragraph text or unrelated package parts.

```ts
import { Document } from 'bun-ooxml';

const document = Document.create();
document.addParagraph('A paragraph with explicit spacing.');
document.paragraphs[0]!.setProperties({
  alignment: 'center', spacingBefore: 120, spacingAfter: 240,
  keepLines: true, pageBreakBefore: false, widowControl: true,
});
await document.save('paragraph.docx');
```

Alignment accepts `left`, `center`, `right` or `both`. Spacing uses integer twips
from 0 through 31680; 20 twips equal one point. The flags accept booleans. For every
field, `null` removes the direct property and `undefined` preserves it. An empty
patch refuses. The getter returns `null` for absent properties, including absent
flags; it does not apply style inheritance or default pagination behaviour.

`outlineLevel` accepts integer OOXML values 0–9. Values 0–8 represent outline
levels 1–9; 9 explicitly means no outline level. An absent property reads `null`,
distinct from explicit `9`. The getter does not infer heading status from a style
ID or resolve inherited outline levels. Outline metadata does not itself change
appearance; consumers may use it for document outlines or tables of contents.
This API does not build a TOC.

A changed edit returns `{ changed: 1 }` and expires held paragraph handles. An
identical request returns `{ changed: 0 }`, retains handles and preserves exact
archive bytes. Both paths still check protection, target freshness and supported
topology. After any successful edit, obtain a fresh paragraph from the document.

Edits require plain direct runs and unique, ordered paragraph properties. Unknown,
revised, duplicated or misordered properties refuse. Selected property attributes
must use the WordprocessingML namespace. Line-based or enabled automatic
before/after spacing refuses because it can override direct twip values. Existing
line spacing and disabled auto-spacing values survive a spacing edit. Rewriting
that spacing leaf can change its prefix, quotes and whitespace; unrelated
paragraph-property fragments and other package members stay untouched.

Changes use the existing paragraph-formatting transaction. A staged write or
serialization failure restores package bytes and leaves the old paragraph handle
usable. String edits preserve an existing UTF-16 encoding through package save.
This operation does not add sections, calculate pagination or change style
inheritance.

## Specification and API policy

ECMA-376 Part 1, fifth edition (October 2016), defines `jc` in §17.3.1.13,
`keepLines` in §17.3.1.14, `pageBreakBefore` in §17.3.1.23, `spacing` in §17.3.1.33
and `widowControl` in §17.3.1.44. Section 17.3.1.20 defines `outlineLvl` and its
0–9 values. The complete PDF is in the
[shared specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).

The standard includes more alignment and spacing forms than this API accepts.
The four alignment values and 31680-twip ceiling are editor policies. In
particular, an absent `widowControl` is not equivalent to an effective false value:
its inherited and default behaviour must be resolved separately. The getter
reports only whether the direct property exists and what it says.

The selected shared scenarios check four alignments, four before/after pairs and
one three-flag combination in memory. Native tests separately check save/reopen,
UTF-16, table custody, rollback and refusal. Outline tests independently save and
reopen all ten values, distinguish absent/removed/explicit non-outline states,
and check lexical no-ops, ordering, malformed metadata, UTF-8 BOMs and both UTF-16
byte orders. They add no canonical heading-style coverage. One generated sample passes Open XML
SDK's Office2019 validation profile. Neither those getter cases nor that schema
check establish Word rendering or pagination parity.
