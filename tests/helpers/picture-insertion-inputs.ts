import assert from 'node:assert/strict';
import {pictureRoot} from './picture-inputs.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';
const manifest=await Bun.file(pictureRoot+'/manifest.json').json();
for(const path of ['workflows/pptx/picture-insertion.feature','contracts/pptx-picture-insertion.md','ledgers/pptx-picture-insertion.json']){
 const asset=manifest.files.find((f:{path:string})=>f.path===path);assert(asset,'registered shared insertion definition');
 const bytes=await Bun.file(pictureRoot+'/'+path).bytes();assert.equal(bytes.length,asset.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),asset.sha256);
}
export const insertionContract=await Bun.file(pictureRoot+'/ledgers/pptx-picture-insertion.json').json();
export const insertionCases=insertionContract.cases;
async function fixture(id:string){const f=manifest.files.find((f:{id:string})=>f.id===id);assert(f?.role==='fixture');const b=await Bun.file(pictureRoot+'/'+f.path).bytes();assert.equal(b.length,f.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(b).digest('hex'),f.sha256);return b;}
export async function insertionInput(c:any){
 const parts=readZip(await fixture(insertionContract.baseFixtureId)),decoder=new TextDecoder(),encoder=new TextEncoder();
 for(const op of c.operations){if(op.kind==='copy-member'){assert(parts.has(op.from)&&!parts.has(op.part));parts.set(op.part,parts.get(op.from)!.slice());}else{assert.equal(op.kind,'replace-literal-once');const s=decoder.decode(parts.get(op.part));assert.equal(s.split(op.before).length,2);parts.set(op.part,encoder.encode(s.replace(op.before,op.after)));}}
 const source=readZip(await fixture(c.request.payload.fixtureId)).get(c.request.payload.part)!;assert.equal(source.length,c.request.payload.byteLength);assert.equal(new Bun.CryptoHasher('sha256').update(source).digest('hex'),c.request.payload.sha256);
 const payload=c.requestPatch?.payloadChange==='empty'?new Uint8Array():source.slice();
 return {bytes:writeZip(parts),source:source.slice(),payload,geometry:{...c.request.geometry,...c.requestPatch?.geometry},options:{...c.request.options,...c.requestPatch?.options}};
}
