import assert from 'node:assert/strict';
import {pictureRoot} from './picture-inputs.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';
export const opacityContract=await Bun.file(pictureRoot+'/ledgers/pptx-opacity.json').json();
export const opacityCases=opacityContract.cases;
const manifest=await Bun.file(pictureRoot+'/manifest.json').json();
for(const path of [opacityContract.feature,opacityContract.contract,'ledgers/pptx-opacity.json']){const asset=manifest.files.find((f:{path:string})=>f.path===path);assert(asset);const bytes=await Bun.file(pictureRoot+'/'+path).bytes();assert.equal(bytes.length,asset.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),asset.sha256);}
export async function opacityInput(c:any){const f=manifest.files.find((f:{id:string})=>f.id===opacityContract.baseFixtureId);assert(f?.role==='fixture');const b=await Bun.file(pictureRoot+'/'+f.path).bytes();assert.equal(b.length,f.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(b).digest('hex'),f.sha256);const parts=readZip(b),decode=new TextDecoder(),encode=new TextEncoder();for(const op of c.operations){assert.equal(op.kind,'replace-literal-once');const s=decode.decode(parts.get(op.part));assert.equal(s.split(op.before).length,2);parts.set(op.part,encode.encode(s.replace(op.before,op.after)));}return {bytes:writeZip(parts),...c.request};}
