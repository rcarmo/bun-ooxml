import {expect} from 'bun:test';
import {Document} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const text='Alpha beta';
export async function formattingDocument(kind='plain'){
 const d=Document.create();d.addParagraph(text);const p=await OpcPackage.open(await d.save());
 let xml=p.text(p.mainPart()).replace('<w:r><w:t>Alpha beta</w:t></w:r>','<w:r><w:rPr><w:b/><w:color w:val="AABBCC"/></w:rPr><w:t xml:space="preserve">Alpha </w:t></w:r><w:r><w:rPr><w:i w:val="0"/><w:u w:val="single"/></w:rPr><w:t>beta</w:t></w:r>');
 if(kind==='field')xml=xml.replace('<w:t>beta</w:t>','<w:instrText>beta</w:instrText>');
 if(kind==='tracked')xml=xml.replace('<w:r><w:rPr><w:i','<w:ins w:id="0" w:author="A"><w:r><w:rPr><w:i').replace('<w:t>beta</w:t></w:r>','<w:t>beta</w:t></w:r></w:ins>');
 if(kind==='mixed-content')xml=xml.replace('</w:r><w:r>','</w:r>unowned<w:r>');
 if(kind==='duplicate-property')xml=xml.replace('<w:b/>','<w:b/><w:b w:val="0"/>');
 if(kind==='malformed-flag')xml=xml.replace('<w:b/>','<w:b w:val="maybe"/>');
 if(kind==='wrong-namespace')xml=xml.replace('<w:b/>','<w:b xmlns:q="urn:foreign" q:val="0"/>');
 if(kind==='property-revision')xml=xml.replace('<w:b/>','<w:b/><w:rPrChange w:id="0" w:author="A"><w:rPr/></w:rPrChange>');
 p.set(p.mainPart(),xml);addPart(p,'custom/format-sentinel.bin',new Uint8Array([9,2,8]),'application/octet-stream');
 if(kind==='protected'){addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');}
 if(kind==='external-settings')addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','https://example.invalid/settings.xml',{external:true});
 return Document.open(p.toBytes());
}
export function directFlags(xml:string){return elements(parseXml(xml),'r',W).map(r=>{const pr=r.children.find(c=>c.localName==='rPr'&&c.namespaceURI===W);const read=(name:string)=>{const node=pr?.children.find(c=>c.localName===name&&c.namespaceURI===W);return node?(attribute(node,'val',W)??'1'):null;};return {bold:read('b'),italic:read('i')};});}
export const bindings:StepBinding[]=[
 {pattern:/^a native Word paragraph split into two differently formatted text runs$/,run:async c=>{c.doc=await formattingDocument();c.before=await(c.doc as Document).save();}},
 {pattern:/^paragraph formatting requests (enable|disable|remove|no-op)$/,run:(c,op)=>{c.op=op;const patch=op==='enable'?{bold:true,italic:true}:op==='disable'?{bold:false,italic:false}:op==='remove'?{bold:null,italic:null}:{};c.result=(c.doc as Document).paragraphs[0]!.setRunFormatting(patch);}},
 {pattern:/^reopening preserves the paragraph text and the (enable|disable|remove|no-op) direct properties$/,run:async(c,op)=>{const bytes=await(c.doc as Document).save(),reopened=await Document.open(bytes);expect(reopened.paragraphs[0]!.text).toBe(text);const flags=directFlags(new TextDecoder().decode(reopened.package.get('word/document.xml')!));expect(flags).toEqual(op==='enable'?[{bold:'1',italic:'1'},{bold:'1',italic:'1'}]:op==='disable'?[{bold:'0',italic:'0'},{bold:'0',italic:'0'}]:op==='remove'?[{bold:null,italic:null},{bold:null,italic:null}]:[{bold:'1',italic:null},{bold:null,italic:'0'}]);expect(c.result).toEqual({changedRuns:op==='no-op'?0:2});}},
 {pattern:/^formatting changes only the main document part and retains unrelated properties$/,run:async c=>{const before=await OpcPackage.open(c.before as Uint8Array),after=await OpcPackage.open(await(c.doc as Document).save());expect(after.names()).toEqual(before.names());for(const name of before.names())if(name!==before.mainPart())expect(after.get(name)).toEqual(before.get(name));expect(after.text(after.mainPart())).toContain('<w:color w:val="AABBCC"/>');expect(after.text(after.mainPart())).toContain('<w:u w:val="single"/>');if(c.op==='no-op')expect(after.toBytes()).toEqual(c.before as Uint8Array);}},
 {pattern:/^a Word formatting refusal input (\S+)$/,run:async(c,kind)=>{const d=await formattingDocument(kind);c.doc=d;c.handle=d.paragraphs[0];if(kind==='stale')d.addParagraph('later');c.before=await d.save();c.kind=kind;}},
 {pattern:/^paragraph direct formatting is attempted$/,run:c=>{try{c.result=(c.handle as any).setRunFormatting(c.kind==='invalid-option'?{bold:'yes'}:{bold:true});}catch(e){c.error=e;}}},
 {pattern:/^formatting refuses before changing package bytes$/,run:async c=>{expect(c.error).toBeDefined();expect((c.error as any).code).toMatch(/^docx-/);expect(await(c.doc as Document).save()).toEqual(Uint8Array.from(c.before as Uint8Array));}},
];
