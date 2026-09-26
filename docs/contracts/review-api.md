# Word stories and tracked revisions

Standalone package APIs avoid mutating Document caches behind live handles. Users
open an OpcPackage, inspect/edit/resolve, save, then reopen Document. Every module
uses Bun-native XML/package primitives without an external producer.

`src/docx/story.ts`:
* `type RevisionView = 'current'|'original'|'all'`
* `inspectStories(pkg:OpcPackage, options?:{view?:RevisionView}): StoryInspection`
* `StoryInspection = {stories:Story[],blindRegions:{part:string;kind:string;count:number}[]}`
* `Story = {part:string;kind:'body'|'header'|'footer'|'footnotes'|'endnotes'|'comments'; paragraphs:{index:number;text:string}[]}`
* `storyParts(pkg): {part:string;kind:Story['kind']}[]` lists linked story parts.
Order body then headers/footers/notes/comments, deterministic by relationship order
then part name. Discover exact officeDocument relationship types from main part
(and referenced story parts) without URL fetch. Report orphan/unsupported story
parts as blind regions if relevant. Namespaces exact; separators in notes skipped.
Current omits deletions, original omits insertions, all includes both. Read p inside
tables/controls supported; unknown/fields/drawings/AlternateContent counted, never
silently asserted complete. Existing Document.paragraphs semantics stay unchanged.

`src/docx/revisions.ts`:
* `inspectRevisions(pkg): {revisions:Revision[];unsupported:RevisionFinding[]}`
* `Revision = {part:string;id:string;kind:'insertion'|'deletion';author?:string;date?:string;text:string}`
* `RevisionFinding = {part:string;kind:string;reason:string}`
* `resolveRevisions(pkg, action:'accept'|'reject', options?:{parts?:string[]}): {resolved:number;changedParts:string[]}`
Bounded run-level w:ins/w:del as direct children of w:p, plain w:r/w:t or w:delText
and rPr safe formatting. All selected stories preflight before mutation; unsupported
move/format/table-row/nested revisions refuse entire scope, as do unsafe namespace
lift, missing/duplicate ids (scope per part), protected docs. No partial resolution.
Accept ins unwrap/del remove; reject ins remove/del unwrap and delText->t preserving
namespace scope and bytes outside selected revision wrappers. No-op untouched bytes.
Inspect unsupported remains read-only. Protection checks read all exact settings
relationships and allow only explicit false enforcement. Missing/unknown enforcement
refuses conservatively. Saving validates inside the transaction before commit.

`src/docx/redline.ts`:
* `trackedReplace(pkg, part:string, query:string, replacement:string, options:{author:string;date:string}): {revisionIds:string[];changedParts:string[]}`
Exactly one occurrence across simple runs within a paragraph; query nonempty,
no tabs/newlines boundaries/fields/drawings, no existing revisions in target story.
Validate author/date, protection, duplicates and content before mutation. Preserve
untouched runs/properties; generate w:del/w:delText and w:ins with author/date/id,
first matched run formatting for insertion. Deletion preserves each matched run's
formatting. Snapshot source XML for stale-scope assumptions only synchronous API.
Before returning, resolve copies accept/reject and assert semantic text matches
replacement/original. Tests check formatting and opaque-part preservation. General
document comparison and paragraph/table change algebra are not implemented.

The root and ./docx entrypoints export these APIs. Author filtering, revision
snapshots, full comparison and other revision types are unsupported. The shared
specification defines the supported story and revision outcomes.
Existing comment inspection and targeted resolution are documented separately in
[comments-api.md](comments-api.md); comment bodies and anchors are never edited.
