# ZIP and XML package admission

`admitPackage(bytes, limits?)` checks archive structure and XML syntax without
requiring `[Content_Types].xml`, an office-document relationship or a complete
OPC graph. It returns a `Map<string, Uint8Array>` of decompressed file payloads.
The input archive is unchanged and returned payloads are detached from it.

```ts
import { admitPackage } from 'bun-ooxml';

const bytes = Uint8Array.from(await Bun.file('input.zip').bytes());
const parts = admitPackage(bytes, {
  maxArchiveBytes: 16 * 1024 * 1024,
  maxEntries: 1000,
  maxEntryBytes: 8 * 1024 * 1024,
  maxTotalBytes: 32 * 1024 * 1024,
  maxCompressionRatio: 100,
});
```

The ZIP reader enforces its existing single-disk ZIP32/ZIP64, name, method,
size and CRC rules. STORED and DEFLATE are accepted; BZIP2 is refused before
decompression. Empty directory entries are omitted. Limit overrides must be
nonnegative safe integers; the compression ratio must be finite and positive.
Defaults are 256 MiB archive bytes, 10,000 entries, 128 MiB per entry, 512 MiB
expanded total and a compression ratio of 1000.

Members ending in `.xml` or `.rels`, ignoring case, are parsed without loading
entities or following relationships. UTF-8 and BOM-prefixed UTF-16LE/BE are
supported. Decoding is strict; an unsupported or contradictory encoding
declaration refuses. Literal member bytes are returned without reserialization.
Admission uses the parser's syntax-only mode, without building descendant text
for every ancestor. DTD declarations, malformed XML and the XML parser's own
depth, node and input limits refuse with `OoxmlError`. ZIP and XML errors retain their existing codes;
invalid limit configuration uses `package-admission-limit-invalid`.

This API neither saves a package nor validates document schemas, relationship
graphs, content-type closure or semantic equivalence. Use `OpcPackage.open` for
OPC graph validation. XML-looking bytes under other suffixes stay opaque.
Decompressed members are collected before XML validation; this API does not
promise streaming, constant memory or an independently measured allocation cap.

The shared ZIP/XML admission cases define rejection outcomes for their exact
inputs. Python's guard uses case-sensitive suffix checks and a compressed-size
precheck against its total budget; Bun's suffix policy, defaults and refusal
ordering differ. Neither admission result establishes cross-runtime parser
parity or package-diff semantics.
