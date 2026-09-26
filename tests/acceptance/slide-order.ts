import {expect} from 'bun:test';
import {Presentation,type Slide} from '../../src/pptx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export function orderFor(kind:string):number[]{return kind==='empty'?[]:kind==='same'?[0,1,2]:kind==='rotate'?[1,2,0]:kind==='duplicate-index'?[0,0,2]:kind==='missing-index'?[0,1]:kind==='out-of-range'?[0,1,3]:kind==='fractional-index'?[0,1,1.5]:[2,1,0];}
export async function orderDeck(kind='reverse'){
 const d=Presentation.create();if(kind==='empty')return d;for(const title of ['Alpha','Beta','Gamma'])d.addTextSlide(title);const p=d.package,main=p.mainPart();let xml=p.text(main);
 if(kind==='duplicate-id'){const ids=elements(parseXml(xml),'sldId',P);xml=xml.replace(`id="${ids[1]!.attributes.id}"`,`id="${ids[0]!.attributes.id}"`);}
 if(kind==='duplicate-list')xml=xml.replace('</p:presentation>','<p:sldIdLst/></p:presentation>');
 if(kind==='lexical-barrier')xml=xml.replace('<p:sldIdLst>','<p:sldIdLst><!--position-->');
 if(kind==='custom-show')xml=xml.replace('</p:presentation>','<p:custShowLst/></p:presentation>');
 if(kind==='extension-metadata')xml=xml.replace('</p:presentation>','<p:extLst/></p:presentation>');
 if(kind==='protected')xml=xml.replace('</p:presentation>','<p:modifyVerifier/></p:presentation>');
 if(kind==='aliased')xml=xml.replaceAll('xmlns:p=','xmlns:q=').replaceAll('xmlns:r=','xmlns:z=').replace(/(<\/?|\s)p:/g,'$1q:').replace(/\sr:/g,' z:');
 if(kind==='default-namespace')xml=xml.replace('xmlns:p=','xmlns=').replace(/<(\/?)p:/g,'<$1');p.set(main,xml);
 if(kind==='duplicate-target'){const rel='ppt/_rels/presentation.xml.rels';p.set(rel,p.text(rel).replace('slides/slide2.xml','slides/slide1.xml'));}
 if(kind==='wrong-mime')p.set('[Content_Types].xml',p.text('[Content_Types].xml').replace('application/vnd.openxmlformats-officedocument.presentationml.slide+xml','application/xml'));
 if(kind==='notes'){for(const [i,s]of d.slides.entries()){const name=`ppt/notesSlides/notesSlide${i+1}.xml`;addPart(p,name,`<p:notes xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Notes ${i}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`,'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');addRelationship(p,s.partName,R+'/notesSlide',`../notesSlides/notesSlide${i+1}.xml`);}}
 return Presentation.open(p.toBytes());
}
export const bindings:StepBinding[]=[
 {pattern:/^a native deck prepared for slide-order (\S+)$/,run:async(c,k)=>{const d=await orderDeck(k);c.doc=d;c.kind=k;c.handles=d.slides;c.before=d.package.toBytes();}},
 {pattern:/^its slides are reordered for (\S+)$/,run:(c,k)=>{c.receipt=(c.doc as Presentation).reorderSlides(orderFor(k!));}},
 {pattern:/^reopened slide order and change receipt match (\S+)$/,run:async(c,k)=>{const d=c.doc as Presentation,q=await Presentation.open(d.package.toBytes()),order=orderFor(k!),titles=['Alpha','Beta','Gamma'];expect(q.slides.map(s=>s.inspectText('order')[0]!.text)).toEqual(order.map(i=>titles[i]!));expect(c.receipt).toEqual({changed:k==='same'||k==='empty'?0:1});if(k==='notes')expect(q.slides.map(s=>s.readNotesText())).toEqual(order.map(i=>`Notes ${i}`));}},
 {pattern:/^slide identities and unrelated package bytes are retained$/,run:async c=>{const d=c.doc as Presentation,before=await OpcPackage.open(c.before as Uint8Array),main=before.mainPart(),order=orderFor(c.kind as string),handles=c.handles as Slide[];expect(d.package.names()).toEqual(before.names());for(const n of before.names())if(n!==main)expect(d.package.get(n)).toEqual(before.get(n));expect(d.slides).toEqual(order.map(i=>handles[i]!));expect(d.slides.map(s=>s.index)).toEqual(order.map((_,i)=>i));const old=before.text(main),now=d.package.text(main),oldList=elements(parseXml(old),'sldIdLst',P)[0]!,newList=elements(parseXml(now),'sldIdLst',P)[0]!;expect(now.slice(0,newList.openEnd)+now.slice(newList.closeStart)).toBe(old.slice(0,oldList.openEnd)+old.slice(oldList.closeStart));expect(newList.children.map(n=>now.slice(n.start,n.end))).toEqual(order.map(i=>old.slice(oldList.children[i]!.start,oldList.children[i]!.end)));if(c.kind==='same'||c.kind==='empty')expect(d.package.toBytes()).toEqual(c.before as Uint8Array);}},
 {pattern:/^an unsafe slide-order input (\S+)$/,run:async(c,k)=>{const d=await orderDeck(k);c.doc=d;c.kind=k;c.handles=d.slides;if(k==='stale-main')d.package.set(d.package.mainPart(),d.package.text(d.package.mainPart())+' ');if(k==='stale-relationships'){const rel='ppt/_rels/presentation.xml.rels';d.package.set(rel,d.package.text(rel)+' ');}c.before=d.package.toBytes();}},
 {pattern:/^its slide permutation is attempted$/,run:c=>{try{c.receipt=(c.doc as Presentation).reorderSlides(orderFor(c.kind as string));}catch(e){c.error=e;}}},
 {pattern:/^slide-order refusal preserves archive bytes and handle order$/,run:c=>{expect(c.error).toBeDefined();expect((c.error as any).code).toMatch(/^PPTX_/);const d=c.doc as Presentation;expect(d.package.toBytes()).toEqual(Uint8Array.from(c.before as Uint8Array));expect(d.slides).toEqual(c.handles as Slide[]);}},
];
