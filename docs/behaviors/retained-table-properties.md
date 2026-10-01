# Retained table properties

The explicit shared candidate `ec227fa96a1d8b00bc4aacf36cb6d51e7fc5d735` adds20 PPTX cases/201 steps and20 Word cases/201 steps. The default pin remains unchanged and shared consumer declarations remain planned.

PPTX table handles expose `patchRetained(row,column,patch)` for cell fills/borders/layout, grid-column width plus frame extent, row height plus frame extent, table flags, frame position and selected cell run bold. Two-span geometry is staged atomically and requires exact original frame/grid/row sums. Successful edits stale the captured table handle; refusals and no-ops leave it usable.

Word table handles expose the same bounded method for cell shading/borders/margins/preferred width/layout/flags, row header/pagination/height and direct table alignment/indentation. Preferred cell width does not rewrite the physical grid. Retained success consumes only the new retained editing target; legacy row/cell getter and setter semantics remain unchanged.

`src/retained-table-xml.ts` contains source-preserving quote-consuming splice helpers. Production paths refuse unsupported ownership, merged/grouped/revised/field/protected/foreign/lexical topology and invalid original/requested values before transaction commit. Exact source profile and bounds are in shared `contracts/retained-table-properties.md`.

`tests/acceptance/retained-table.ts` independently parses saved member XML, reopens, checks exact values/child order, masks only edited tokens/children or the two dimension attributes, and compares every other byte/member/graph. Word margin comparison preserves the literal unpatched type token rather than dropping each entire margin leaf. Native controls include wrong types/late invalid patches, executable getters, namespaces/quoted decoys, malformed source, merge/protection/tracking, stale handles, staged rollback, save-failure destinations and oracle corruption. Production no-op faults in six families and neighbor-text corruption fail assertions and pass after exact restoration.

The candidate full lane adds40 to the earlier cumulative acceptance total (852 implemented,59 planned). Exact selected receipt is40/402. Historical identity helpers remove only sealed additive IDs/predicates for older checks. Inventory and source-seal mapping updates confer no execution credit or support percentage.

No Office rendering, theme/gradient/alpha border styling, table creation/deletion/merging/import, publication, push/tag or default-pin advancement is part of this batch.
