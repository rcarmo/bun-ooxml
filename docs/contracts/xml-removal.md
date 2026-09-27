# XML subtree removal

`XmlSnapshot.parse(source)` creates a syntax-checked, immutable XML string
snapshot. Its `elements` array contains source-ordered handles with expanded
names and UTF-16 source offsets. `remove(targets)` returns a new string with the
selected subtrees removed. All other source characters keep their exact spelling,
including comments, declarations, whitespace, namespace prefixes and quotes.

```ts
import { XmlSnapshot } from 'bun-ooxml';

const snapshot = XmlSnapshot.parse('<r><!--keep--><drop/><keep/></r>');
const targets = snapshot.elements.filter(t => t.localName === 'drop');
const result = snapshot.remove(targets);
// <r><!--keep--><keep/></r>
```

Handles belong to one snapshot. Copying a handle or supplying one from another
snapshot raises `XML_REMOVAL_TARGET`, even when both source strings are equal.
The root cannot be removed (`XML_REMOVAL_ROOT`). Duplicate selections and
ancestor/descendant selections raise `XML_REMOVAL_OVERLAP`. Input order does not
matter. An empty selection returns the original string, and every operation
leaves the snapshot available for independent subsequent operations.

Both the source and result pass the native XML syntax and namespace checks.
Removing an element can join text into an illegal sequence, such as `]]>`; this
raises `XML_REMOVAL_UNSAFE` without returning output. Parsing enforces the existing
limits of 8 Mi UTF-16 code units, 256 levels and 100,000 elements. Snapshot parsing
and output validation do not accumulate descendant text at every ancestor.

This API accepts JavaScript strings; [XmlByteSnapshot](xml-byte-snapshot.md)
provides owned UTF-8/UTF-16 byte input and encoding-preserving removal, attribute
updates and structured-content edits. Neither API saves package members,
validates an OOXML schema, updates relationships or resolves references to deleted
content. Use the format-level editors when those checks are needed. Parse a new
snapshot to edit the returned string further; original handles remain tied to
the original source. The same handles also support
[source-preserving attribute updates](xml-attributes.md) and
[structured insertion/replacement](xml-structure.md).
