import assert from 'node:assert/strict';
import {pictureRoot} from './picture-inputs.ts';
import {insertionInput} from './picture-insertion-inputs.ts';
export const placementContract=await Bun.file(pictureRoot+'/ledgers/pptx-picture-placement.json').json();
export const placementCases=placementContract.cases;
export async function placementInput(c:any){
 const input=await insertionInput({operations:[],request:{payload:c.request.payload,geometry:c.request.box,options:{contentType:c.request.options.contentType}}});
 return {...input,box:{...c.request.box,...c.requestPatch?.box},options:{...c.request.options,...c.requestPatch?.options}};
}
async function verifyPlacementDefinitions(){const manifest=await Bun.file(pictureRoot+'/manifest.json').json();for(const path of [placementContract.feature,placementContract.contract,'ledgers/pptx-picture-placement.json']){const asset=manifest.files.find((f:{path:string})=>f.path===path);assert(asset);const b=await Bun.file(pictureRoot+'/'+path).bytes();assert.equal(b.length,asset.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(b).digest('hex'),asset.sha256);}}
await verifyPlacementDefinitions();
