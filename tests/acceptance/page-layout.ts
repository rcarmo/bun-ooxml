import {expect} from 'bun:test';
import {Document,type PageLayout,W_NS as W} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const originalLayout:PageLayout={width:12240,height:15840,orientation:'portrait',top:1440,right:1800,bottom:1440,left:1800,header:720,footer:720,gutter:0};
export function desired(kind:string):PageLayout{return {...originalLayout,...(kind==='same'?{}:kind==='landscape'?{width:15840,height:12240,orientation:'landscape' as const}:kind==='margins'?{top:1000,left:900,right:900,bottom:1000,header:400,footer:400,gutter:120}:{width:11906,height:16838}),...(kind==='invalid-width'?{width:0}:kind==='negative-margin'?{left:-1}:kind==='no-content-area'?{left:9000,right:9000}:kind==='invalid-orientation'?{orientation:'sideways' as any}:{})};}
export async function layoutDocument(kind='portrait'){
 const d=Document.create();d.addParagraph('Layout text π');const p=await OpcPackage.open(await d.save());let xml=p.text(p.mainPart());
 if(kind==='missing-section')xml=xml.replace(/<w:sectPr>[\s\S]*?<\/w:sectPr>/,'');
 if(kind==='duplicate-section')xml=xml.replace('</w:body>','<w:sectPr/></w:body>');
 if(kind==='misplaced-section')xml=xml.replace('</w:body>','<w:p/></w:body>');
 if(kind==='duplicate-size')xml=xml.replace('<w:pgSz','<w:pgSz w:w="10" w:h="20"/><w:pgSz');
 if(kind==='missing-margins')xml=xml.replace(/<w:pgMar[^>]*\/>/,'');
 if(kind==='wrong-namespace')xml=xml.replace('w:w="12240"','w="12240"');
 if(kind==='section-revision')xml=xml.replace('</w:sectPr>','<w:sectPrChange w:id="1" w:author="Prior"><w:sectPr/></w:sectPrChange></w:sectPr>');
 if(kind==='lexical-barrier')xml=xml.replace('<w:pgSz','<!-- barrier --><w:pgSz');
 if(kind==='earlier-section')xml=xml.replace(/(<w:p\b[^>]*>)/,'$1<w:pPr><w:sectPr><w:pgSz w:w="10000" w:h="14000"/></w:sectPr></w:pPr>');
 if(kind==='header-reference'){addPart(p,'word/header1.xml',`<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Keep header</w:t></w:r></w:p></w:hdr>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');const rel=addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header','header1.xml');xml=xml.replace('<w:sectPr>',`<w:sectPr><w:headerReference w:type="default" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rel.id}"/>`);}
 if(kind==='aliased')xml=xml.replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:');
 if(kind==='default-namespace')xml=xml.replace('xmlns:w=', 'xmlns=').replace(/<(\/?)w:/g,'<$1').replace('<document ','<document xmlns:w="'+W+'" ');
 p.set(p.mainPart(),xml);
 if(kind==='protected'){addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');}
 addPart(p,'custom/layout.bin',new Uint8Array([3,4,5]),'application/octet-stream');return Document.open(p.toBytes());
}
function withoutGeometry(xml:string){const doc=parseXml(xml),body=doc.root.children.find(n=>n.localName==='body'&&n.namespaceURI===W)!,section=body.children.find(n=>n.localName==='sectPr'&&n.namespaceURI===W)!;const targets=section.children.filter(n=>n.namespaceURI===W&&['pgSz','pgMar'].includes(n.localName)).sort((a,b)=>b.start-a.start);for(const node of targets)xml=xml.slice(0,node.start)+xml.slice(node.end);return xml;}
export const bindings:StepBinding[]=[
 {pattern:/^a native document prepared for page-layout (\S+)$/,run:async(c,k)=>{c.kind=k;c.doc=await layoutDocument(k);c.before=await(c.doc as Document).save();}},
 {pattern:/^the final section page layout is selected for (\S+)$/,run:(c,k)=>{c.receipt=(c.doc as Document).setPageLayout(desired(k!));}},
 {pattern:/^saved and reopened page geometry matches (\S+)$/,run:async(c,k)=>{const q=await Document.open(await(c.doc as Document).save());expect(q.getPageLayout()).toEqual(desired(k!));expect(c.receipt).toEqual({changed:k==='same'?0:1});expect(q.paragraphs[0]!.text).toBe('Layout text π');const xml=new TextDecoder().decode(q.package.get('word/document.xml')!),body=parseXml(xml).root.children.find(n=>n.localName==='body')!,section=body.children.find(n=>n.localName==='sectPr')!,size=section.children.find(n=>n.localName==='pgSz')!;expect(Number(attribute(size,'w',W))).toBe(desired(k!).width);expect(Number(attribute(size,'h',W))).toBe(desired(k!).height);}},
 {pattern:/^other section properties, earlier sections and package payloads are unchanged$/,run:async c=>{const before=await OpcPackage.open(c.before as Uint8Array),after=await OpcPackage.open(await(c.doc as Document).save());expect(after.names()).toEqual(before.names());for(const name of before.names())if(name!==before.mainPart())expect(after.get(name)).toEqual(before.get(name));expect(withoutGeometry(after.text(after.mainPart()))).toBe(withoutGeometry(before.text(before.mainPart())));if(c.kind==='same')expect(after.toBytes()).toEqual(c.before as Uint8Array);}},
 {pattern:/^an unsafe final-section layout input (\S+)$/,run:async(c,k)=>{const d=await layoutDocument(k);c.doc=d;c.kind=k;c.handle=d.paragraphs[0];if(k==='stale-document')d.package.setPart('word/document.xml',new TextEncoder().encode(new TextDecoder().decode(d.package.get('word/document.xml')!).replace('Layout text','External text')));c.before=await d.save();}},
 {pattern:/^its page-layout change is attempted$/,run:c=>{try{c.receipt=(c.doc as Document).setPageLayout(desired(c.kind as string));}catch(e){c.error=e;}}},
 {pattern:/^page-layout selection refuses without changing archive bytes or handles$/,run:async c=>{expect(c.error).toBeDefined();expect((c.error as any).code).toMatch(/^docx-/);const d=c.doc as Document;expect(await d.save()).toEqual(Uint8Array.from(c.before as Uint8Array));expect(d.paragraphs[0]).toBe(c.handle as Document['paragraphs'][number]);}},
];
