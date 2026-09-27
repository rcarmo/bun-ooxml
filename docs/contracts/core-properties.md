# Word document core properties

`Document.getCoreProperties()` returns detached direct metadata values.
`setCoreProperties(patch)` updates supplied fields and returns the number of
changed fields as `{ changed: number }`. It creates the core-properties part,
content-type entry and package relationship when needed.

```ts
document.setCoreProperties({
  title: 'Quarterly report',
  creator: 'Example author',
  modified: '2026-02-03T01:00:00Z',
});
const title = document.getCoreProperties().title;
```

The fields are `title`, `creator`, `subject`, `description`, `identifier`,
`language`, `keywords`, `category`, `contentStatus`, `lastModifiedBy`, `revision`,
`version`, `created`, `modified` and `lastPrinted`. Every field is a string or
`null`. Null means absent on read and removes the field in a patch. Undefined
leaves it unchanged. Empty strings are distinct from absent values, except that
timestamp fields require a date. Revision and version remain strings.

Each string is limited to 4096 UTF-16 code units and valid XML characters.
Timestamps use full UTC seconds: `YYYY-MM-DDTHH:mm:ssZ`, with valid calendar dates
and years 0001–9999. Reduced dates, offsets, fractions and leap seconds refuse.
This is a bounded editor profile; the specification admits additional date forms.
No dates, authors or revision counters are updated automatically.

## Preservation and refusals

The package must have at most one core-properties relationship and one part with
the core-properties media type. The relationship must be internal and originate
at the package root. Orphan core-property parts, duplicate targets, wrong types,
invalid roots and an occupied default creation path refuse. Creation uses
`docProps/core.xml`; existing valid relationships can name a different part.

The reader and writer check every field, even for no-ops or removal. Unknown,
duplicate or misqualified elements, decorated fields, nested content and lexical
barriers refuse. Created and modified require an `xsi:type` QName resolving to
`dcterms:W3CDTF`. Last-printed uses plain `cp:lastPrinted`. Namespace aliases are
accepted; newly authored fields declare their own namespace bindings.

Identical decoded values preserve exact archive bytes, including entity spelling.
Changed edits retain unselected field fragments and unrelated package members.
The metadata write preserves an existing UTF-8 BOM or UTF-16 endianness/BOM.
Removing all fields retains the empty metadata part and relationship rather than
silently deleting package structure. An all-null patch on a document without
metadata does not create a part.

Plain patch fields are copied without invoking accessors. Creation and updates
run in a package transaction, including content-type and relationship writes,
readback and serialization. Failure rolls back metadata and retains document
handles. Successful metadata edits also retain paragraph/table handles because
the main document does not change. The main-part snapshot must still match;
out-of-band main-part changes refuse. Metadata itself is read from the current
package on each call. Protected documents permit reads but refuse writes,
including same-value writes, under this editor's protection policy.

The XML parser/editor's complete 8 Mi UTF-16-unit limit applies to existing and
edited metadata. Input-field and XML limits are not peak-memory guarantees.

## Specification and independent checks

ECMA-376 Part 2, fifth edition (December 2021), §§8.2–8.3 specifies core-properties
parts and markup; Annex E lists the media type, relationship and namespaces.
See the [complete specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).
The shared scenario supplies all fields but compares only title, creator and
subject in memory. Native tests separately check all 15 fields after disk
save/reopen, rollback, namespace aliases and encoding custody. The Office oracle
reads all 15 authored values through Open XML SDK's package-properties API, in
addition to its package validation. Automatic metadata changes by Office and
full date-format interoperability have not been tested.
