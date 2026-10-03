import assert from 'node:assert/strict';
import {resolve} from 'node:path';
import type {PictureInfo} from '../../src/pptx/index.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';

// Explicit local-development candidate; the released reference gitlink stays unchanged.
export const pictureRoot=resolve(process.env.OOXML_GRAPHICS_ROOT??resolve(import.meta.dir,'../../../fixtures-ooxml'));
export type PictureRecipe={caseId:string;scenarioId:string;example?:{fault:string};operations:{kind:string;part:string;before:string;after:string}[];expected?:PictureInfo[];errorCode?:string};
const manifest=await Bun.file(pictureRoot+'/manifest.json').json();
for(const path of ['workflows/pptx/pictures.feature','contracts/pptx-pictures.md','ledgers/pptx-picture-inspection.json']){
 const asset=manifest.files.find((f:{path:string})=>f.path===path);assert(asset,'registered shared picture definition');
 const bytes=await Bun.file(pictureRoot+'/'+path).bytes();assert.equal(bytes.length,asset.bytes);
 assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),asset.sha256);
}
export const pictureContract=await Bun.file(pictureRoot+'/ledgers/pptx-picture-inspection.json').json();
export const pictureRecipes:PictureRecipe[]=pictureContract.cases;
const fixture=manifest.files.find((f:{id:string})=>f.id===pictureContract.baseFixtureId);
assert(fixture?.role==='fixture','known canonical picture fixture');
export async function pictureInput(recipe:PictureRecipe):Promise<Uint8Array>{
 const bytes=await Bun.file(pictureRoot+'/'+fixture.path).bytes();
 assert.equal(bytes.length,fixture.bytes);
 assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),fixture.sha256);
 const parts=readZip(bytes),encoder=new TextEncoder(),decoder=new TextDecoder();
 for(const op of recipe.operations){
  assert.equal(op.kind,'replace-literal-once');assert(parts.has(op.part),'recipe member exists');
  const text=decoder.decode(parts.get(op.part));assert.equal(text.split(op.before).length,2,'literal recipe occurrence');
  parts.set(op.part,encoder.encode(text.replace(op.before,op.after)));
 }
 return recipe.operations.length?writeZip(parts):bytes;
}
