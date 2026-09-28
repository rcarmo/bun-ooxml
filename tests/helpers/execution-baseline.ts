import assert from 'node:assert/strict';
const added=['persistence','custody','no-op','refusal','rollback','plain-edit','author-refusal'].map(n=>'@id-docx-tracking-settings-'+n);
const mergeIds=['roundtrip','content-refusal','structure-refusal','coordinate-refusal','rollback','encoding','stale'].map(n=>'@id-docx-horizontal-merge-'+n);
/** Keep the v0.27 execution identities stable while checking each reviewed addition. */
export function withoutTrackingToggle(keys:string[]):string[]{
 assert.equal(keys.length,730);
 const chain:string[]=keys.filter(k=>k==='@id-xlsx-owned-calculation-chain-invalidation');assert.equal(chain.length,1);
 const descriptor:string[]=keys.filter(k=>k==='@id-zip-unsigned-descriptor-signature-collision');
 assert.equal(descriptor.length,1);
 const admission=keys.filter(k=>k.startsWith('@id-package-admission-negative-budget|'));
 assert.equal(admission.length,2);assert.equal(new Set(admission).size,2);
 const moves=keys.filter(k=>k.startsWith('@id-docx-paired-move-'));assert.equal(moves.length,45);assert.equal(new Set(moves).size,45);
 const properties=keys.filter(k=>k.startsWith('@id-docx-run-property-revisions-'));assert.equal(properties.length,33);assert.equal(new Set(properties).size,33);
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
 return keys.filter(k=>!chain.includes(k)&&!descriptor.includes(k)&&!admission.includes(k)&&k!=='@id-docx-go-track-author-toggle'&&k!=='@id-docx-go-table-cell-access'&&!settings.includes(k)&&!merges.includes(k)&&!verticals.includes(k)&&!templates.includes(k)&&!threads.includes(k)&&!properties.includes(k)&&!moves.includes(k));
}
