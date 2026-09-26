# Archive payload comparison

`comparePackageArchives(before, after, limits?)` compares admitted ZIP file
members without requiring an OPC graph. It returns sorted lists named `added`,
`removed`, `changed`, `equivalent_xml` and `unchanged`. Each member name appears
in exactly one list. Inputs are byte arrays; the operation writes no files and
does not modify either archive.

```ts
import { comparePackageArchives } from 'bun-ooxml';

const before = Uint8Array.from(await Bun.file('original.docx').bytes());
const after = Uint8Array.from(await Bun.file('edited.docx').bytes());
const report = comparePackageArchives(before, after, { maxTotalBytes: 32 * 1024 * 1024 });
```

Both archives independently pass [`admitPackage`](package-admission.md) with the
same limits before any report is returned. Invalid ZIP structure, duplicate
names, CRC failures, unsafe XML or exceeded budgets throw. This applies even to
identical, added or removed XML members. The budgets apply per archive; both
expanded maps coexist during comparison, so they are not a combined process
memory limit.

Byte-identical common payloads are `unchanged`. Other common `.xml` and `.rels`
members, ignoring suffix case, use [`xmlEquivalent`](xml-comparison.md): true
places the name in `equivalent_xml`, false in `changed`. All other changed
payloads are binary. Empty directories are omitted; ZIP ordering, compression,
timestamps and archive metadata do not affect these payload categories. Equality
of bytes does not establish schema validity or XML semantic support.

This API does not resolve content types, relationships, signatures, rendered
appearance or application meaning. Its XML category inherits the comparator's
conservative and schema-limited policy. It reports member names only, without
payload hashes. The existing `diffPackages()` API still compares bytes and
resolved content types: prefix-only XML changes remain `changed` there.

The shared semantic-diff scenario checks its exact prefix-only XML, binary
replacement and addition. Python's helper does not admit archives first, uses
case-sensitive XML suffixes and also returns hashes for changed payloads. Those
behaviours are not claimed by this Bun API. Duplicate archives and invalid XML
are refused here rather than compared under Python's looser input policy.
