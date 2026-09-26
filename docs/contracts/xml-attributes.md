# XML attribute updates

`XmlSnapshot.setAttributes(patches)` updates or adds attributes without rewriting
surrounding source characters. Each patch supplies an issued element handle, a
lexical QName and a string value:

```ts
import { XmlSnapshot } from 'bun-ooxml';

const snapshot = XmlSnapshot.parse('<r xmlns:p="urn:p"><t a = \'old\'/></r>');
const target = snapshot.elements[1]!;
const result = snapshot.setAttributes([
  { target, name: 'a', value: "x'y&z" },
  { target, name: 'p:flag', value: 'yes' },
]);
// <r xmlns:p="urn:p"><t a = 'x&#39;y&amp;z' p:flag="yes"/></r>
```

Names resolve in the target's original namespace scope. An unprefixed attribute
has no namespace, even under a default element namespace. A prefix alias selects
an existing attribute by its expanded name and preserves the attribute's original
QName, quotes and spacing. The `xml` prefix works without a declaration. An absent
attribute uses the supplied QName, double quotes and one leading space before the
original start-tag closing delimiter.

Values are escaped for the existing quote style. Tabs, carriage returns and line
feeds use character references so they survive XML attribute normalisation.
Assigning the current decoded value leaves the original spelling untouched. An
empty batch returns the original string. Root and descendant attributes can be
changed in one batch; two patches for the same expanded attribute on one element
refuse, including aliases and identical values.

Handles belong to the issuing snapshot; copied or foreign handles refuse.
Namespace declarations, unbound prefixes, malformed names and invalid XML
characters also refuse. Validation finishes before returning any output, and the
original snapshot stays available after either success or refusal. Parse a new
snapshot to edit the returned string further.

Batches have at most 100,000 patches. Aggregate input name/value lengths and
incrementally planned output length are limited to 8 Mi UTF-16 code units. These
are string limits, not UTF-8 byte limits; a later shrinking patch does not rescue
an earlier over-limit intermediate plan. Output passes the same XML syntax checks
as [subtree removal](xml-removal.md), without accumulating descendant text at each
ancestor. Namespace scope walks are bounded by the parser's depth limit; no
peak-memory or throughput guarantee is established.

This operation edits XML strings. It does not remove attributes, create namespace
bindings, validate an OOXML schema, update references or save package members.
