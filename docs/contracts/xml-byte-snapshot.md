# XML byte snapshots

`XmlByteSnapshot.parse(bytes)` copies a `Uint8Array` into an immutable snapshot.
`remove(targets)` returns a new, detached byte array with selected XML subtrees
removed. Both caller input and the snapshot remain unchanged.

```ts
import { XmlByteSnapshot } from 'bun-ooxml';

const input = await Bun.file('part.xml').bytes();
const snapshot = XmlByteSnapshot.parse(input);
const output = snapshot.remove(
  snapshot.elements.filter(node => node.localName === 'drop'),
);
```

The frozen `elements` array contains the same kind of handles as
[string snapshots](xml-removal.md): expanded names and offsets measured in decoded
UTF-16 code units, excluding the byte-order mark. These are not byte offsets.
Handles belong to their issuing snapshot. String-snapshot handles, foreign or
copied handles, root removal and overlapping selections refuse.

## Encodings and custody

Supported input is strict UTF-8, optionally with a BOM, or UTF-16LE/BE with a BOM.
A declaration, if present, must agree with the detected encoding; generic UTF-16
accepts either UTF-16 BOM. Unsupported codecs, BOM-less UTF-16, malformed byte
sequences, duplicate leading BOMs and conflicting declarations refuse. No
replacement characters are introduced during decoding.

An empty removal returns an exact copy of the original bytes. Changed output
retains the original encoding, BOM, declaration and spelling of all surviving
source characters. Re-encoding does not normalise entities, whitespace, namespace
prefixes or quotes. Mutating the caller buffer or either an empty or changed
result cannot change subsequent edits from the snapshot.

Subview boundaries are honoured, including Node-compatible Buffer inputs.
Caller overrides of `slice`, iterators, `byteLength` and `buffer` do not control
admission or copying. SharedArrayBuffer-backed and detached inputs refuse. The
API accepts byte views only; wrap an owned ArrayBuffer in a Uint8Array first.

The byte limit is 32 MiB, checked before copying. Decoded XML uses the existing
8 Mi UTF-16-unit, 256-level and 100,000-element scanner limits. These limits do
not bound peak process memory: parsing retains owned bytes, decoded text and
snapshot metadata, and editing allocates output.

Every removal validates the joined XML, including text that would become an
illegal `]]>` sequence. The original snapshot can be reused after success or
failure. This byte API currently exposes subtree removal only; attribute and
structured-content edits remain on the string API. It does not read/write OPC
packages, validate OOXML schemas or repair relationships to removed content.

The shared byte-input case checks one UTF-8 parse/no-op and caller-byte custody.
Native tests separately cover changed removal, caller/result mutation, both UTF-16
endiannesses, BOMs, malformed inputs and bounds. The binding executes Bun even
though the historical scenario text names its Go API profile; no Go execution or
cross-runtime equivalence is inferred.
