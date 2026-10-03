import assert from 'node:assert/strict';
import {pictureRoot} from './picture-inputs.ts';
const manifest=await Bun.file(pictureRoot+'/manifest.json').json();
export const officeSmartArtContract=await Bun.file(pictureRoot+'/ledgers/pptx-smartart-office-source.json').json();
for(const path of [officeSmartArtContract.feature,officeSmartArtContract.contract,'ledgers/pptx-smartart-office-source.json']){
 const asset=manifest.files.find((f:{path:string})=>f.path===path);assert(asset);const bytes=await Bun.file(pictureRoot+'/'+path).bytes();assert.equal(bytes.length,asset.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),asset.sha256);
}
export async function officeSmartArtInput(){const fixture=manifest.files.find((f:{id:string})=>f.id===officeSmartArtContract.fixtureId);assert(fixture?.role==='fixture');const bytes=await Bun.file(pictureRoot+'/'+fixture.path).bytes();assert.equal(bytes.length,fixture.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),fixture.sha256);return bytes;}
