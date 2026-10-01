import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {OpcPackage} from '../../src/opc/package.ts';
import {inspectTextShapes,patchShapeText,setShapeAutofit} from '../../src/pptx/shape-text.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
const A='http://schemas.openxmlformats.org/drawingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main';
const deck=()=>{const d=Presentation.create();d.addTextSlide('Original','Subtitle');return d;};
const title=(d:Presentation)=>d.slides[0]!.inspectTextShapes().find(s=>s.placeholder==='ctrTitle')!;
test('shape text retains Unicode, XML values, CR, whitespace and empty LF paragraphs on saved reopen',async()=>{
 const d=deck(),id=title(d).shapeId,part=d.slides[0]!.partName,before=d.package.text(part),others=d.slides[0]!.inspectTextShapes().filter(s=>s.shapeId!==id);
 d.slides[0]!.setShapeText(id,'\n snow 雪 & <value> 😀\r\n\nend\r');
 const q=await Presentation.open(d.package.toBytes());expect(title(q).text).toBe('\n snow 雪 & <value> 😀\n\nend\r');expect(q.slides[0]!.inspectTextShapes().filter(s=>s.shapeId!==id)).toEqual(others);
 const old=parseXml(before),now=parseXml(q.package.text(part)),oldShape=elements(old,'sp',P)[0]!,newShape=elements(now,'sp',P)[0]!;expect(q.package.text(part).slice(0,newShape.start)).toBe(before.slice(0,oldShape.start));expect(q.package.text(part).slice(newShape.end)).toBe(before.slice(oldShape.end));
});
test('append and clear preserve exact bodyPr and lstStyle; detached inspections cannot mutate live text',()=>{
 const d=deck(),s=d.slides[0]!,id=title(d).shapeId,part=s.partName;const original=d.package.text(part);const pr=elements(parseXml(original),'bodyPr',A)[0]!,lst=elements(parseXml(original),'lstStyle',A)[0]!;
 s.setShapeText(id,'Second',true);expect(title(d).text).toBe('Original\nSecond');const values=s.inspectTextShapes();values[0]!.paragraphs[0]!.text='bad';expect(title(d).text).toBe('Original\nSecond');s.clearShapeText(id);expect(title(d).text).toBe('');const updated=d.package.text(part);expect(updated).toContain(original.slice(pr.start,pr.end));expect(updated).toContain(original.slice(lst.start,lst.end));
});
test('bullet levels and label flags are exact, invalid arguments refuse atomically without anchor version change',()=>{
 const d=deck(),s=d.slides[0]!,id=title(d).shapeId;s.clearShapeText(id);s.addBullet(id,'& value',8,'Label');expect(title(d).text).toBe('Label: & value');expect(title(d).paragraphs[0]!.level).toBe(8);
 const before=d.package.toBytes(),version=d.currentSlideVersion(s.partName);for(const operation of [()=>s.addBullet(id,'bad',9),()=>s.addBullet(id,'bad',-1),()=>s.addBullet(id,'a\nb'),()=>s.setShapeText(id,'\u0000'),()=>s.setShapeText(999,'bad'),()=>s.setShapeAutofit(id,'toString' as any),()=>s.setShapeText(id,'bad',1 as any)]){expect(operation).toThrow();expect(d.package.toBytes()).toEqual(before);expect(d.currentSlideVersion(s.partName)).toBe(version);}
});
test('autofit preserves existing attributes and unrelated shape properties, replaces exactly one fit child',()=>{
 const d=deck(),id=title(d).shapeId,part=d.slides[0]!.partName;let xml=d.package.text(part).replace('<a:bodyPr/>','<a:bodyPr wrap="square" anchor="ctr"><a:noAutofit/></a:bodyPr>');let next=setShapeAutofit(xml,id,'shrink');expect(next).toContain('wrap="square" anchor="ctr"');expect(elements(parseXml(next),'noAutofit',A)).toHaveLength(0);expect(elements(parseXml(next),'normAutofit',A)).toHaveLength(1);next=setShapeAutofit(next,id,'resize');expect(elements(parseXml(next),'spAutoFit',A)).toHaveLength(1);expect(inspectTextShapes(next)[0]!.text).toBe('Original');
 expect(()=>setShapeAutofit(xml.replace('<a:noAutofit/>','<a:noAutofit/><a:spAutoFit/>'),id,'none')).toThrow();
});
test('aliases preserve custody, rich fields/comments/duplicate IDs and text locks refuse before a replacement result',()=>{
 const d=deck(),id=title(d).shapeId,xml=d.package.text(d.slides[0]!.partName);const alias=xml.replaceAll('xmlns:a=','xmlns:z=').replace(/(<\/?|\s)a:/g,'$1z:');const changed=patchShapeText(alias,id,'New');expect(inspectTextShapes(changed)[0]!.text).toBe('New');
 const mutations=[xml.replace('<a:r><a:t>Original</a:t></a:r>','<a:fld id="field"><a:t>Original</a:t></a:fld>'),xml.replace('<a:r><a:t>Original</a:t></a:r>','<!--keep--><a:r><a:t>Original</a:t></a:r>'),xml.replace('<a:t>Original</a:t>','<a:t>Original<!--keep--></a:t>'),xml.replace('<a:t>Original</a:t>','<a:t><a:x/>Original</a:t>'),xml.replace('noGrp="1"','noTextEdit="1"'),xml.replace('id="3"','id="2"')];for(const source of mutations)expect(()=>patchShapeText(source,id,'Changed')).toThrow();
});
test('shape edit invalidates text/table anchors and production serialization failure rolls back package and version',()=>{
 const d=deck(),s=d.slides[0]!,id=title(d).shapeId;s.addTable(1,1,{x:0,y:0,width:100,height:100});const table=s.tables[0]!,anchor=s.inspectText('before')[0]!.anchor;s.setShapeText(id,'Changed');expect(()=>table.cell(0,0)).toThrow('stale');expect(()=>s.replaceTextAt(anchor,'Original','bad')).toThrow('stale');
 const before=d.package.toBytes(),version=d.currentSlideVersion(s.partName),original=OpcPackage.prototype.toBytes;OpcPackage.prototype.toBytes=function(){throw Error('injected serialization');};try{expect(()=>s.setShapeText(id,'Lost')).toThrow('injected');}finally{OpcPackage.prototype.toBytes=original;}expect(d.package.toBytes()).toEqual(before);expect(d.currentSlideVersion(s.partName)).toBe(version);expect(title(d).text).toBe('Changed');
});
test('positioned insertion preserves held slide identities, handles empty decks and rolls back all parts/metadata on serialization fault',async()=>{
 const d=deck(),held=d.slides[0]!,anchor=held.inspectText('before')[0]!.anchor;d.insertTextSlide(0,'First','subtitle');expect(held.index).toBe(1);held.replaceTextAt(anchor,'Original','Updated');expect((await Presentation.open(d.package.toBytes())).slides.map(s=>s.inspectText('saved')[0]!.text)).toEqual(['First','Updated']);const empty=Presentation.create();expect(empty.insertTextSlide(0,'only').index).toBe(0);
 const before=d.package.toBytes(),handles=d.slides,original=OpcPackage.prototype.toBytes;let calls=0;OpcPackage.prototype.toBytes=function(){calls++;if(calls>=2)throw Error('injected insertion');return original.call(this);};try{expect(()=>d.insertTextSlide(1,'bad')).toThrow('injected');}finally{OpcPackage.prototype.toBytes=original;}expect(d.package.toBytes()).toEqual(before);expect(d.slides).toEqual(handles);expect(held.index).toBe(1);for(const n of [-1,2.5,99]){expect(()=>d.insertTextSlide(n,'bad')).toThrow();expect(d.package.toBytes()).toEqual(before);}
});
test('shape edit retains UTF16 BOM/declaration and source archive, only one decoded payload changes',async()=>{
 for(const le of [true,false]){const d=deck(),s=d.slides[0]!,part=s.partName,xml=d.package.text(part).replace('encoding="UTF-8"','encoding="UTF-16"'),bytes=new Uint8Array(2+xml.length*2),v=new DataView(bytes.buffer);v.setUint16(0,0xfeff,le);for(let i=0;i<xml.length;i++)v.setUint16(2+i*2,xml.charCodeAt(i),le);d.package.set(part,bytes);const source=d.package.toBytes(),q=await Presentation.open(source),old=new Map(q.package.names().map(n=>[n,q.package.get(n)!]));q.slides[0]!.setShapeText(title(q).shapeId,'雪');const saved=await Presentation.open(q.package.toBytes());expect([...saved.package.get(part)!.slice(0,2)]).toEqual(le?[255,254]:[254,255]);expect(saved.package.text(part)).toContain('encoding="UTF-16"');expect(title(saved).text).toBe('雪');for(const[n,b]of old)if(n!==part)expect(saved.package.get(n)).toEqual(b);expect((await Presentation.open(source)).package.toBytes()).toEqual(source);}
});
