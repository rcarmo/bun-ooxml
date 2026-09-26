# ZIP64 contract

`src/opc/zip.ts` supports bounded ZIP64 read/write for single-disk archives.

## Read support

`readZip(bytes, limits?)` accepts ZIP64 only when all metadata can be validated without guessing:

* single-disk only; any non-zero disk number still fails as `zip-multi-disk-unsupported`
* ZIP64 EOCD locator must be present immediately before the classic EOCD
* ZIP64 EOCD record must be the fixed 44-byte body; extensible data sectors are refused as `zip-zip64-unsupported`
* all ZIP64 64-bit integers must fit `Number.MAX_SAFE_INTEGER`; otherwise `zip-zip64-unsupported`
* the declared central-directory count must fit within the declared central-directory span before iteration
* the central directory ends exactly at the end-record boundary; undeclared gaps refuse
* central-directory entries with sentinel `0xffff` / `0xffffffff` fields must carry exactly one ZIP64 extra field (`0x0001`) with exactly the required values in the required order
* duplicate or ambiguous ZIP64 extra fields are refused as `zip-zip64-unsupported`
* missing or truncated required ZIP64 extra data is malformed and fails as:
  * `zip-structure-invalid` for central-directory metadata
  * `zip-local-metadata-mismatch` for local-header metadata
* ZIP64 local headers and ZIP64 data descriptors must agree with the central directory; disagreements fail as `zip-local-metadata-mismatch`

Existing ZIP safety limits are unchanged:

* `zip-archive-too-large`
* `zip-too-many-entries`
* `zip-entry-too-large`
* `zip-total-too-large`
* `zip-compression-ratio-exceeded`
* CRC validation still fails as `zip-crc-mismatch`
* inflate output is capped by the declared size, with a one-byte cap for empty streams followed by exact zero-length verification; size disagreements fail as `zip-size-mismatch`

## Write support

`writeZip(parts, options?)` now accepts:

```ts
writeZip(parts, { forceZip64?: boolean })
```

Rules:

* ZIP64 is written automatically for `>= 65535` entries
* ZIP64 is also written when any member size, compressed size, local-header offset, central-directory size or central-directory offset reaches the reserved ZIP32 sentinel value
* `forceZip64: true` emits tiny valid ZIP64 output for tests, even when sizes fit ZIP32
* output remains single-disk only
* member names still must fit the ZIP 16-bit name length field; overlong names fail as `zip-zip64-unsupported`
* archives that cannot be materialized with safe integer offsets/lengths fail as `zip-zip64-unsupported`

The writer materialises the entire archive in memory. Forced tiny archives,
65,535-entry archives and resource/preflight refusals are tested. Multi-gigabyte
writes and independent Office producer/consumer checks have not been run.

## Error-code summary

* malformed/missing ZIP64 locator or EOCD: `zip-structure-invalid`
* unsupported ZIP64 extensible sector: `zip-zip64-unsupported`
* ZIP64 values beyond safe integer range: `zip-zip64-unsupported`
* ambiguous/duplicate ZIP64 extras: `zip-zip64-unsupported`
* missing/truncated central ZIP64 extras: `zip-structure-invalid`
* missing/truncated local ZIP64 extras or local ZIP64 disagreement: `zip-local-metadata-mismatch`
* multi-disk ZIP64: `zip-multi-disk-unsupported`
