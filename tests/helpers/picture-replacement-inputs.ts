import assert from 'node:assert/strict';
import {pictureRoot} from './picture-inputs.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';
const manifest=await Bun.file(pictureRoot+'/manifest.json').json();
for(const path of ['workflows/pptx/picture-replacement.feature','contracts/pptx-picture-replacement.md','ledgers/pptx-picture-replacement.json']){
 const asset=manifest.files.find((f:{path:string})=>f.path===path);assert(asset,'registered shared replacement definition');const bytes=await Bun.file(pictureRoot+'/'+path).bytes();assert.equal(bytes.length,asset.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),asset.sha256);
}
export const replacementContract=await Bun.file(pictureRoot+'/ledgers/pptx-picture-replacement.json').json();
export const replacementCases=replacementContract.cases;
export async function replacementInput(c:any){
 const f=manifest.files.find((f:{id:string})=>f.id===replacementContract.baseFixtureId);assert(f?.role==='fixture');const bytes=await Bun.file(pictureRoot+'/'+f.path).bytes();assert.equal(bytes.length,f.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),f.sha256);
 const original=readZip(bytes),parts=new Map(original),decode=new TextDecoder(),encode=new TextEncoder();
 for(const op of c.operations){assert.equal(op.kind,'replace-literal-once');const s=decode.decode(parts.get(op.part));assert.equal(s.split(op.before).length,2);parts.set(op.part,encode.encode(s.replace(op.before,op.after)));}
 const source=original.get(c.request.payload.part)!;assert.equal(c.request.payload.fixtureId,f.id);assert.equal(source.length,c.request.payload.byteLength);assert.equal(new Bun.CryptoHasher('sha256').update(source).digest('hex'),c.request.payload.sha256);
 return {bytes:writeZip(parts),payload:source.slice(),source:source.slice(),shapeId:c.requestPatch?.shapeId??c.request.shapeId,options:{...c.request.options,...c.requestPatch?.options}};
}
