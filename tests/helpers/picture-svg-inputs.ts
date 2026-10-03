import assert from 'node:assert/strict';
import {pictureRoot} from './picture-inputs.ts';
import {insertionInput} from './picture-insertion-inputs.ts';
export const svgContract=await Bun.file(pictureRoot+'/ledgers/pptx-picture-svg.json').json();
export const svgCases=svgContract.cases;
const manifest=await Bun.file(pictureRoot+'/manifest.json').json();for(const path of [svgContract.feature,svgContract.contract,'ledgers/pptx-picture-svg.json']){const asset=manifest.files.find((f:{path:string})=>f.path===path);assert(asset);const bytes=await Bun.file(pictureRoot+'/'+path).bytes();assert.equal(bytes.length,asset.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),asset.sha256);}
export async function svgInput(c:any){const input=await insertionInput({operations:c.operations,request:{payload:c.request.fallback,geometry:c.request.geometry,options:c.request.options}});return {...input,svg:new TextEncoder().encode(c.request.svg),fallback:c.requestPatch?.emptyFallback?new Uint8Array():input.payload};}
