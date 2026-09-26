# Existing Word comments

`inspectComments(pkg)` reads existing comments and reply links without creating or
rewriting parts. `setCommentResolved(pkg, id, resolved)` changes only the selected
existing `commentEx` resolution flag. Both functions are exported from the root
and `./docx` entrypoints and take an `OpcPackage` snapshot.

```ts
type CommentInfo = {
  id: string;
  author?: string;
  initials?: string;
  date?: string;
  text: string;
  paragraphId?: string;
  parentId?: string;
  resolved: boolean;
};
type CommentInspection = {
  comments: CommentInfo[];
  unsupported: { part: string; kind: string }[];
};

inspectComments(pkg): CommentInspection;
setCommentResolved(pkg, id: string, resolved: boolean): {
  changed: number;
  changedParts: string[];
};
```

## Inspection

Exact relationships and content types identify the linked comments and extension
parts. Plain paragraphs/runs, tabs and line breaks are read; paragraphs join with
LF. Inspection returns detached values. Malformed IDs, duplicate paragraph keys,
missing reply parents, cycles and invalid done values refuse. Unsupported bodies
and unlinked comment parts are reported and prevent resolution edits.

A comment without matching extended metadata can be inspected, but cannot be
resolved through this API. Missing done means unresolved. Valid existing values
are `0`, `1`, `true`, `false`, `on` and `off`. IDs are normalised decimal integers;
paragraph IDs are eight hexadecimal digits. Modern commentsIds/extensible/person
metadata is preserved without interpretation; authoring or deletion of those
graphs is unsupported.

## Resolution

Resolution requires a supported comment graph and an existing linked entry. It
changes no body, anchor, reply parent or other comment's state. An existing done
attribute retains its QName, quote style and surrounding whitespace; only its
value becomes `1` or `0`. A missing flag is added with the namespace prefix of the
existing paragraph-ID attribute, including under a default element namespace.
The part's supported original encoding/BOM is retained.

A same-state request preserves exact archive bytes and returns zero changes.
Otherwise the receipt names only the extension part. The edit, readback and
serialization validation run inside one synchronous rollback boundary. Refusal
or serialization failure retains all earlier package edits unchanged. Protected
settings refuse unless enforcement is explicitly disabled; every linked settings
part is checked, and external or invalid settings refuse.

Creation, deletion, body rewriting, anchor validation/repair and thread-wide
resolution are outside this API. Run it before opening a format reader or reopen
readers afterward; it does not refresh a live `Document` cache.

## Checks and limits

`features/docx/comments.feature` binds 14 observable cases using the canonical
comments fixture. Native tests add namespace/lexical/encoding custody, absent flags,
protected settings, detached inspection, disk save/reopen and injected rollback.
No shared fixture is modified. CommentsExtended MIME uses the vendor-backed
OpenXML value recorded in the shared fact registry; the disputed alias refuses.
Native XML/package readback passed; independent Word reopening/rendering has not
been performed. An independent review attempt timed out and supplied no evidence.
