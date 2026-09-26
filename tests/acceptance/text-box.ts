import {expect} from 'bun:test';
import {Presentation,type TextBoxGeometry,type TextBoxOptions} from '../../src/pptx/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main';
export const geometry:TextBoxGeometry={x:120,y:240,width:914400,height:457200};
export const texts:Record<string,string>={plain:'Box & <π>',multiline:' first\r\n\nlast\r',empty:'',formatted:'Direct flags'};
export async function boxDeck(kind:string){
 const d=Presentation.create();d.addTextSlide('Existing title');d.addTextSlide('Other slide');const s=d.slides[0]!,p=d.package;let xml=p.text(s.partName);
 if(kind==='alias')xml=xml.replaceAll('xmlns:p=','xmlns:q=').replaceAll('xmlns:a=','xmlns:d=').replace(/(<\/?|\s)p:/g,'$1q:').replace(/(<\/?|\s)a:/g,'$1d:');
 if(kind==='default-namespace')xml=xml.replaceAll('<p:','<').replaceAll('</p:','</').replace('xmlns:p=', 'xmlns=');
 if(kind==='extension-tail')xml=xml.replace('</p:spTree>','<p:extLst><p:ext uri="opaque"><q:data xmlns:q="urn:custom">keep</q:data></p:ext></p:extLst></p:spTree>');
 if(kind==='nested-id')xml=xml.replace('</p:spTree>','<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="95" name="Nested shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></p:grpSp></p:spTree>');
 if(kind==='duplicate-id')xml=xml.replace('id="2"','id="1"');
 if(kind==='malformed-id')xml=xml.replace('id="2"','id="2junk"');
 if(kind==='exhausted-id')xml=xml.replace('id="2"','id="2147483647"');
 if(kind==='duplicate-tree')xml=xml.replace('</p:cSld>','<p:spTree/></p:cSld>');
 if(kind==='missing-prefix-properties')xml=xml.replace(/<p:nvGrpSpPr>[\s\S]*?<\/p:nvGrpSpPr>/,'');
 if(kind==='misplaced-extension')xml=xml.replace('<p:sp>','<p:extLst/><p:sp>');
 if(kind==='alternate-content')xml=xml.replace('</p:spTree>','<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/></p:spTree>');
 if(kind==='transformed-tree')xml=xml.replace('<a:off x="0" y="0"/>','<a:off x="10" y="0"/>');
 p.set(s.partName,xml);
 if(kind==='protected-presentation'){const main=p.mainPart();p.set(main,p.text(main).replace('</p:presentation>','<p:modifyVerifier cryptProviderType="rsaAES"/></p:presentation>'));}
 return d;
}
export function request(kind:string):{text:string;geometry:TextBoxGeometry;options:TextBoxOptions}{
 return {text:kind==='invalid-text'?'bad\u0001':texts[kind]??texts.plain!,geometry:{...geometry,...(kind==='negative-position'?{x:-1}:kind==='zero-extent'?{height:0}:kind==='fractional-geometry'?{x:1.5}:kind==='oversized-geometry'?{width:2147483648}:{})},options:kind==='invalid-options'?{bold:'yes'} as any:kind==='formatted'?{name:'Box & "Name"',bold:true,italic:false}:{}};
}
export const bindings:StepBinding[]=[
 {pattern:/^a native slide prepared for text-box (\S+)$/,run:async(c,k)=>{c.doc=await boxDeck(k!);c.before=(c.doc as Presentation).package.toBytes();}},
 {pattern:/^a positioned text box is appended for (\S+)$/,run:(c,k)=>{const r=request(k!);c.receipt=(c.doc as Presentation).slides[0]!.addTextBox(r.text,r.geometry,r.options);}},
 {pattern:/^reopened text-box identity geometry and paragraphs match (\S+)$/,run:async(c,k)=>{
  const d=await Presentation.open((c.doc as Presentation).package.toBytes()),s=d.slides[0]!,receipt=c.receipt as any,r=request(k!),xml=d.package.text(s.partName),root=parseXml(xml),def=elements(root,'sp',P).find(n=>elements(n,'cNvPr',P).some(id=>id.attributes.id===String(receipt.shapeId)))!;expect(def).toBeDefined();expect(receipt).toEqual({shapeId:k==='nested-id'?96:3,partName:s.partName,paragraphCount:r.text.replace(/\r\n?/g,'\n').split('\n').length});
  const id=elements(def,'cNvPr',P)[0]!;expect(id.attributes.name).toBe(r.options.name??`TextBox ${receipt.shapeId}`);expect(elements(def,'cNvSpPr',P)[0]!.attributes.txBox).toBe('1');expect(elements(def,'off',A)[0]!.attributes).toEqual({x:'120',y:'240'});expect(elements(def,'ext',A)[0]!.attributes).toEqual({cx:'914400',cy:'457200'});expect(elements(def,'noAutofit',A)).toHaveLength(1);expect(elements(def,'t',A).map(t=>t.text)).toEqual(r.text.replace(/\r\n?/g,'\n').split('\n'));expect(s.inspectText('box').slice(1).map(p=>p.text)).toEqual(r.text.replace(/\r\n?/g,'\n').split('\n'));expect(elements(def,'rPr',A)).toHaveLength(k==='formatted'?1:0);for(const run of elements(def,'rPr',A)){expect(run.attributes.b).toBe(r.options.bold===undefined?undefined:r.options.bold?'1':'0');expect(run.attributes.i).toBe(r.options.italic===undefined?undefined:r.options.italic?'1':'0');}
 }},
 {pattern:/^existing slide shapes and unrelated package payloads are unchanged$/,run:async c=>{const d=c.doc as Presentation,before=await Presentation.open(c.before as Uint8Array),part=d.slides[0]!.partName;expect(d.package.names()).toEqual(before.package.names());for(const n of before.package.names())if(n!==part)expect(d.package.get(n)).toEqual(before.package.get(n));const old=before.package.text(part),after=d.package.text(part),tree=elements(parseXml(after),'spTree',P)[0]!,inserted=tree.children.filter(n=>n.namespaceURI===P&&n.localName==='sp'&&elements(n,'cNvPr',P).some(id=>id.attributes.id===String((c.receipt as any).shapeId)));expect(inserted).toHaveLength(1);expect(after.slice(0,inserted[0]!.start)+after.slice(inserted[0]!.end)).toBe(old);}},
 {pattern:/^an unsafe text-box authoring input (\S+)$/,run:async(c,k)=>{c.kind=k;c.doc=await boxDeck(k!);const d=c.doc as Presentation;c.before=d.package.toBytes();c.version=d.currentSlideVersion(d.slides[0]!.partName);}},
 {pattern:/^the unsafe positioned text box is attempted$/,run:c=>{const r=request(c.kind as string);try{c.receipt=(c.doc as Presentation).slides[0]!.addTextBox(r.text,r.geometry,r.options);}catch(e){c.error=e;}}},
 {pattern:/^text-box creation refuses before package bytes or slide version change$/,run:c=>{expect(c.error).toBeDefined();expect((c.error as any).code).toMatch(/^(PPTX_|XML_)/);const d=c.doc as Presentation;expect(d.package.toBytes()).toEqual(Uint8Array.from(c.before as Uint8Array));expect(d.currentSlideVersion(d.slides[0]!.partName)).toBe(c.version as number);}},
];
