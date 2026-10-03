import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {officeSmartArtInput} from '../helpers/smartart-office-inputs.ts';
import {readZip} from '../../src/opc/zip.ts';
import {parseXml,elements,applyEdits} from '../../src/xml/index.ts';
const DSP='http://schemas.microsoft.com/office/drawing/2008/diagram',A='http://schemas.openxmlformats.org/drawingml/2006/main';
test('native Office SmartArt refuses ambiguous slide/data drawing owners before edits',async()=>{
 const source=await Presentation.open(await officeSmartArtInput()),s=source.slides[0]!;
 source.package.set('ppt/diagrams/_rels/data1.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/></Relationships>');
 const destination=Presentation.create(),to=destination.addTextSlide('Refusal'),before=readZip(destination.package.toBytes()),version=destination.currentSlideVersion(to.partName);expect(()=>s.inspectSmartArt()).toThrow(expect.objectContaining({code:'PPTX_SMARTART_UNSUPPORTED'}));expect(()=>to.copySmartArtFrom(s,4)).toThrow(expect.objectContaining({code:'PPTX_SMARTART_UNSUPPORTED'}));const after=readZip(destination.package.toBytes());expect([...after.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)expect(after.get(n)).toEqual(b);expect(destination.currentSlideVersion(to.partName)).toBe(version);
});
test('native Office SmartArt refuses referenced repeated zeros and duplicate nonzero drawing IDs atomically',async()=>{
 for(const variant of ['referenced-zero','duplicate-nonzero']){
  const source=await Presentation.open(await officeSmartArtInput()),path='ppt/diagrams/drawing1.xml',xml=source.package.text(path),doc=parseXml(xml),target=Presentation.create(),to=target.addTextSlide('Refusal'),before=readZip(target.package.toBytes()),version=target.currentSlideVersion(to.partName);
  if(variant==='referenced-zero'){const pr=elements(doc,'cNvSpPr',DSP)[0]!;source.package.set(path,applyEdits(xml,[{start:pr.start,end:pr.end,value:`<dsp:cNvSpPr><a:stCxn xmlns:a="${A}" id="0" idx="0"/></dsp:cNvSpPr>`}]));}
  else source.package.set(path,xml.replaceAll('id="0"','id="5"'));
  expect(()=>to.copySmartArtFrom(source.slides[0]!,4)).toThrow(expect.objectContaining({code:'PPTX_SMARTART_UNSUPPORTED'}));const after=readZip(target.package.toBytes());expect([...after.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)expect(after.get(n)).toEqual(b);expect(target.currentSlideVersion(to.partName)).toBe(version);
 }
});
test('native Office SmartArt metadata allocation rolls back a late output failure',async()=>{
 const source=await Presentation.open(await officeSmartArtInput()),target=Presentation.create(),to=target.addTextSlide('Rollback'),before=readZip(target.package.toBytes()),version=target.currentSlideVersion(to.partName),serialize=target.package.toBytes;
 target.package.toBytes=()=>{if(target.package.text(to.partName).includes('relIds'))throw Error('Office metadata serialization failed');return serialize.call(target.package);};try{expect(()=>to.copySmartArtFrom(source.slides[0]!,4)).toThrow('Office metadata serialization failed');}finally{target.package.toBytes=serialize;}
 const after=readZip(target.package.toBytes());expect([...after.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)expect(after.get(n)).toEqual(b);expect(target.currentSlideVersion(to.partName)).toBe(version);expect(to.copySmartArtFrom(source.slides[0]!,4).shapeId).toBe(3);
});
