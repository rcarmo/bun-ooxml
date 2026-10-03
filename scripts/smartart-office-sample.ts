/** @script Prepare sealed Office SmartArt source and isolated copies for external checks.
 * @usage bun scripts/smartart-office-sample.ts [output-directory]
 * @description Reads a manifest-sealed candidate fixture; creates oracle artifacts only.
 */
import {mkdir} from 'node:fs/promises';
import {resolve} from 'node:path';
import {Presentation} from '../src/pptx/index.ts';
import {officeSmartArtInput,officeSmartArtContract} from '../tests/helpers/smartart-office-inputs.ts';
const out=resolve(process.argv[2]??resolve(import.meta.dir,'../artifacts/smartart-uno'));await mkdir(out,{recursive:true});
const bytes=await officeSmartArtInput();await Bun.write(out+'/source.pptx',bytes);
const copies=[];
for(const same of [false,true]){
 const source=await Presentation.open(bytes),from=source.slides[0]!,target=same?source:Presentation.create(),to=same?from:target.addTextSlide('Copied Office SmartArt');
 const receipt=to.copySmartArtFrom(from,officeSmartArtContract.shapeId),file=(same?'same-slide-copy':'cross-presentation-copy')+'.pptx';await target.save(out+'/'+file);copies.push({file,receipt});
}
await Bun.write(out+'/inputs.json',JSON.stringify({sourceFixtureId:officeSmartArtContract.fixtureId,copies,limits:['No SmartArt editing or layout evaluation','No Microsoft PowerPoint application test']},null,2)+'\n');
console.log('Prepared manifest-sealed SmartArt source and two copies: '+out);
