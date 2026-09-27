import assert from 'node:assert/strict';
/** Only the reviewed tracking-toggle activation may extend the v0.27 baseline. */
export function withoutTrackingToggle(keys:string[]):string[]{
 assert.equal(keys.length,526);
 assert.equal(keys.filter(k=>k==='@id-docx-go-track-author-toggle').length,1);
 return keys.filter(k=>k!=='@id-docx-go-track-author-toggle');
}
