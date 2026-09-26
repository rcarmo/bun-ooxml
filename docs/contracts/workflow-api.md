# Native workflow API (batch 7)

`src/workflow/index.ts` exports `patchOffice(request): Promise<PatchReceipt>`.
A request is `{ source: string, output?: string, format?: 'docx'|'pptx'|'xlsx',
mode: 'dry_run'|'strict'|'safe', changes: {target: string, value: string|number|boolean|null}[],
multilineWrap?: boolean, calculationPolicy?: 'invalidate-without-recalculation',
expectedSourceSha256?: string, expectedDestinationSha256?: string|null }`.

All modes resolve targets against one private source snapshot. This first workflow
requires all targets to be unique, supported and non-overlapping. It refuses the
entire batch otherwise, including safe mode; best-effort is not implemented.
Dry run may stage in memory but never writes. Safe requires a distinct destination;
strict defaults to source but accepts a distinct output. Never silently fall back.

Receipt fields: `status: 'preview'|'refused'|'committed'`, `committedChanges: number`,
`results: {target,value,matched:number,status:'matched'|'unmatched'|'unsupported'|'unchanged'|'committed',code?:string}[]`,
`sourceSha256`, `outputSha256?: string`, `calculationState:'unchanged'|'recalculation-required'`,
`error?: {code:string,message:string}`, `changedParts:string[]`.
Preview reports zero committed operations, actual per-target matches and requested
values. Committed counts come from changed staged operations after successful
atomic save, never from input length. Refusal has zero and no committed results.

Targets: DOCX exact literal body/table text; one match only in this first workflow.
PPTX `slide:N/title` or `slide:N/subtitle`, 1-based slide number, unique real placeholder
shape and one supported paragraph. XLSX `Sheet!A1` or `A1` on the first sheet, existing
cell only. Null clearing and formula-cell overwrites refuse. No formula assignment
via value; strings beginning `=` are literal text. Complex DOCX topology outside
the supported body/table subset refuses the workflow instead of hiding blind regions.

Guarded helper APIs:
* `src/xlsx/styles.ts`: `setCellWrapText(workbook: Workbook, sheetName: string, address: string, enabled: boolean): void`.
  Uses existing styles, clones target xf and alignment before assignment, validates
  style dependencies, qualifies XML names, refuses unsupported topology atomically.
* `src/pptx/placeholders.ts`: `findPlaceholderText(slide: Slide, kind:'title'|'subtitle'): {text:string; anchor:TextAnchor}`.
  Resolves `p:ph type=title|ctrTitle|subTitle` in direct shapes, must be unique,
  one supported text paragraph; no heuristic first paragraph selection.

Calculation receipts describe worksheet `<f>` cell cache invalidation only. Array
and data-table result ranges refuse value edits before mutation because followers
may lack `<f>`. Chart/external-link/opaque caches and calculation chains are
preserved; no general cache freshness or calculation certification is implied.

The workflow implementation owns async file identity checks and serialises its own
writers by canonical source/output path. It does not claim isolation against an
external filesystem writer between final check and rename. Final symlinks and
safe-output hardlink aliases refuse. Caller directories must be trusted.

Contract tests bind the frozen shared v2 feature with only lifecycle-tag changes;
source text, scenario IDs, expanded examples and arguments remain pinned. Native
library support and workflow coverage are separate from inherited OOXML API parity.
