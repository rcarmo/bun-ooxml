# OPC graph edit batch

`src/opc/graph.ts` exports:

* `addPart(pkg, name, value: Uint8Array|string, contentType): void`
* `removePart(pkg, name): void`
* `addRelationship(pkg, owner: string, type: string, target: string, options?: { id?: string; external?: boolean }): Relationship`
* `removeRelationship(pkg, owner: string, id: string): void`
* `nextPartName(pkg, template: string): string` with exactly one `%d`, starting at 1.
* `walkParts(pkg, owner?: string): string[]` follows internal relationships in order, cycle-safe, excludes external targets.

Empty owner is package root. Existing valid packages only. Each mutation runs
inside `pkg.transaction` and validates graph/content-type output before returning.
Part creation refuses collisions; deletion refuses incoming relationships and
main-part deletion. Deleting a part removes its own .rels and content-type
Overrides but does not recursively delete target parts. Relationship removal
refuses Office relationship-id/embed/link references in owner XML; non-XML owners
with outgoing relationships refuse rather than assuming reference semantics.
Owner XML does not get patched automatically. These helpers operate on OpcPackage;
format readers cache some model state, so callers must reopen the format after
graph edits. No cascade, import or arbitrary
relationship URI rewriting in this slice.

`src/opc/content-types.ts` exports:
* `getContentType(pkg, name): string | undefined`
* `setPartContentType(pkg, name, type): void` adds or updates an exact Override,
  unless current default/override already yields this type; preserve untouched XML.
* `removePartContentType(pkg, name): void` removes only matching Override.
These helpers are synchronous within a graph transaction; no disk writes. Default
extensions remain untouched because other parts may depend on them. Root and
child namespaces must be validated; MIME values must have token/token plus optional
parameters or refuse (conservative exact type/subtype sufficient initially).

`src/opc/diff.ts` exports:
* `diffPackages(before: OpcPackage, after: OpcPackage): PackageDiffReport`
* `{added:string[],removed:string[],changed:string[],unchanged:string[],parts: PartChange[]}`
* PartChange: `{name,kind:'added'|'removed'|'changed',beforeSha256?:string,afterSha256?:string,beforeBytes?:number,afterBytes?:number,beforeContentType?:string,afterContentType?:string}`
Report is payload/length/hash/type-based and deterministic. It does not claim XML
semantic equivalence, visual equality, formula correctness or archive-byte equality.
Opaque parts remain opaque. A content-type-only change must mark the affected part
changed even when its payload is equal, as well as the content-type XML member.
