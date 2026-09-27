import assert from 'node:assert/strict';
const added=['persistence','custody','no-op','refusal','rollback','plain-edit','author-refusal'].map(n=>'@id-docx-tracking-settings-'+n);
const mergeIds=['roundtrip','content-refusal','structure-refusal','coordinate-refusal','rollback','encoding','stale'].map(n=>'@id-docx-horizontal-merge-'+n);
/** Keep the v0.27 execution identities stable while checking each reviewed addition. */
export function withoutTrackingToggle(keys:string[]):string[]{
 assert.equal(keys.length,648);
 const threads=keys.filter(k=>k.startsWith('@id-docx-existing-thread-'));assert.equal(threads.length,23);assert.equal(new Set(threads).size,23);
 const templateIds=['values','empty','placeholders','refusal','bounds','encoding','snapshot','scope'].map(n=>'@id-docx-template-inventory-'+n),templates=keys.filter(k=>templateIds.some(id=>k===id||k.startsWith(id+'|')));
 assert.equal(templates.length,22);assert.equal(new Set(templates).size,22);
 const verticals=keys.filter(k=>mergeIds.some(id=>k===id.replace('horizontal','vertical')||k.startsWith(id.replace('horizontal','vertical')+'|')));
 assert.equal(verticals.length,26);assert.equal(new Set(verticals).size,26);
 const merges=keys.filter(k=>mergeIds.some(id=>k===id||k.startsWith(id+'|')));
 assert.equal(merges.length,26);assert.equal(new Set(merges).size,26);
 assert.equal(keys.filter(k=>k==='@id-docx-go-table-cell-access').length,1);
 assert.equal(keys.filter(k=>k==='@id-docx-go-track-author-toggle').length,1);
 const settings=keys.filter(k=>added.some(id=>k===id||k.startsWith(id+'|')));
 assert.equal(settings.length,24);assert.equal(new Set(settings).size,24);
 return keys.filter(k=>k!=='@id-docx-go-track-author-toggle'&&k!=='@id-docx-go-table-cell-access'&&!settings.includes(k)&&!merges.includes(k)&&!verticals.includes(k)&&!templates.includes(k)&&!threads.includes(k));
}
