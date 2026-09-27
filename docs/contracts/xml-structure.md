# Structured XML insertion and replacement

`XmlSnapshot.appendChildren(patches)` appends authored content to existing
elements. `replaceElements(patches)` replaces disjoint non-root subtrees. Each
patch uses a handle issued by that snapshot and a `children` array containing
strings or structured elements:

```ts
import { XmlSnapshot } from 'bun-ooxml';

const snapshot = XmlSnapshot.parse('<r xmlns="urn:root"><slot/><keep/></r>');
const result = snapshot.appendChildren([{
  target: snapshot.elements[1]!,
  children: [{
    name: { localName: 'item', namespaceURI: 'urn:item' },
    attributes: [{ name: { localName: 'id', namespaceURI: '' }, value: '42' }],
    children: ['Text & more'],
  }],
}]);
```

Authored element and attribute names use `{ localName, namespaceURI }`; local
names cannot contain a prefix. The renderer reuses in-scope prefixes or allocates
an unused `n1`, `n2`, etc. An element can use the current default namespace, while
a namespaced attribute always needs a prefix. An empty-namespace element clears
an inherited default namespace with `xmlns=""`. The XML namespace uses `xml`.
Siblings receive independent scopes, and replacement uses the surviving parent's
bindings rather than bindings on the removed subtree.

Text and attribute values are escaped. Attribute whitespace and text carriage
returns use character references to survive reparsing. Authored namespace
declarations and duplicate expanded attributes refuse. Comments, processing
instructions and raw XML fragments are not accepted as authored content; existing
ones outside the edited ranges retain their source spelling.

Appending to a self-closing element preserves its attributes and spacing, replaces
`/>` with `>`, and adds a matching closing tag. Other insertions leave existing
child content intact and append before the original closing tag. Replacements
remove only their selected ranges. An empty replacement payload removes that
subtree; an empty insertion payload leaves the element unchanged. An empty batch
returns the original string.

Both operations refuse foreign or copied handles, duplicate targets and
ancestor/descendant targets, even when their payloads are empty. Root replacement
refuses; appending to the root is supported. The original snapshot remains usable
after success or refusal. Output is checked for XML syntax, including illegal
text sequences created by joining previously separated content.

## Limits

Batches and each array have at most 100,000 entries. The renderer permits at most
100,000 authored content/attribute visits, counts existing ancestors toward the
256-level depth limit, and rejects cycles. Aggregate input fields and final output
are bounded to 8 Mi UTF-16 code units. Escaped output growth is counted before
replacement strings are allocated. Namespace scope processing has a separate
limit of 1,000,000 entry visits per batch. These are bounded API policies; they do
not establish peak-memory or throughput guarantees. The final document also
passes the existing parser's element-count limit.

`XmlSnapshot` operates on JavaScript strings. [XmlByteSnapshot](xml-byte-snapshot.md)
exposes the same operations on owned UTF-8/UTF-16 input and returns detached bytes
in the original encoding. Neither API validates OOXML schemas, resolves references,
reconciles package relationships or saves files. The shared namespace case exercises
100 combinations within one acceptance case; it does not represent 100 independent
specification requirements. The separate byte-input seed executes parse/no-op custody
only; changed byte-structure edits have native tests.
