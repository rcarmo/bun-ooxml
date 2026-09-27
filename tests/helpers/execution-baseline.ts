import assert from 'node:assert/strict';
const added=['persistence','custody','no-op','refusal','rollback','plain-edit','author-refusal'].map(n=>'@id-docx-tracking-settings-'+n);
const mergeIds=['roundtrip','content-refusal','structure-refusal','coordinate-refusal','rollback','encoding','stale'].map(n=>'@id-docx-horizontal-merge-'+n);
/** Only the reviewed tracking, nullable-cell and physical-merge cases extend v0.27. */
export function withoutTrackingToggle(keys:string[]):string[]{
 assert.equal(keys.length,577);
 const merges=keys.filter(k=>mergeIds.some(id=>k===id||k.startsWith(id+'|')));
 assert.equal(merges.length,26);assert.equal(new Set(merges).size,26);
 assert.equal(keys.filter(k=>k==='@id-docx-go-table-cell-access').length,1);
 assert.equal(keys.filter(k=>k==='@id-docx-go-track-author-toggle').length,1);
 const settings=keys.filter(k=>added.some(id=>k===id||k.startsWith(id+'|')));
 assert.equal(settings.length,24);assert.equal(new Set(settings).size,24);
 return keys.filter(k=>k!=='@id-docx-go-track-author-toggle'&&k!=='@id-docx-go-table-cell-access'&&!settings.includes(k)&&!merges.includes(k));
}
