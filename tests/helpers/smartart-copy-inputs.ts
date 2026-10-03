import assert from 'node:assert/strict';
import {pictureRoot} from './picture-inputs.ts';
import {smartArtInput} from './smartart-inputs.ts';
export const copyContract=await Bun.file(pictureRoot+'/ledgers/pptx-smartart-copy.json').json();
export const copyCases=copyContract.cases;
const manifest=await Bun.file(pictureRoot+'/manifest.json').json();for(const path of [copyContract.feature,copyContract.contract,'ledgers/pptx-smartart-copy.json']){const asset=manifest.files.find((f:{path:string})=>f.path===path);assert(asset);const bytes=await Bun.file(pictureRoot+'/'+path).bytes();assert.equal(bytes.length,asset.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),asset.sha256);}
export const copyInput=smartArtInput;
