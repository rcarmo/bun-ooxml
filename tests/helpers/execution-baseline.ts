import assert from 'node:assert/strict';
const added=['persistence','custody','no-op','refusal','rollback','plain-edit','author-refusal'].map(n=>'@id-docx-tracking-settings-'+n);
/** Only the reviewed tracking and nullable-cell cases may extend v0.27. */
export function withoutTrackingToggle(keys:string[]):string[]{
 assert.equal(keys.length,551);
 assert.equal(keys.filter(k=>k==='@id-docx-go-table-cell-access').length,1);
 assert.equal(keys.filter(k=>k==='@id-docx-go-track-author-toggle').length,1);
 const settings=keys.filter(k=>added.some(id=>k===id||k.startsWith(id+'|')));
 assert.equal(settings.length,24);assert.equal(new Set(settings).size,24);
 return keys.filter(k=>k!=='@id-docx-go-track-author-toggle'&&k!=='@id-docx-go-table-cell-access'&&!settings.includes(k));
}
