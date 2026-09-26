import {expect} from 'bun:test';
import {Document} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const STYLE_TYPE='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
export const STYLE_REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
export async function styleDocument(kind='assign'){
 const d=Document.create();d.addParagraph('Styled π text',{bold:true});const p=await OpcPackage.open(await d.save());
 const style=`<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading One"/></w:style>`;
 let styles=`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Body"/>${style}<w:style w:type="character" w:styleId="Emphasis"/></w:styles>`;
 if(kind==='duplicate-style')styles=styles.replace('</w:styles>',style+'</w:styles>');
 addPart(p,'word/styles.xml',styles,kind==='wrong-mime'?'application/xml':STYLE_TYPE);addRelationship(p,p.mainPart(),STYLE_REL,'styles.xml');
 if(kind==='duplicate-relationship'){addPart(p,'word/styles2.xml',styles,STYLE_TYPE);addRelationship(p,p.mainPart(),STYLE_REL,'styles2.xml');}
 if(kind==='external-styles'){const rels='word/_rels/document.xml.rels';p.set(rels,p.text(rels).replace('Target="styles.xml"','Target="https://example.invalid/styles.xml" TargetMode="External"'));}
 let props=['replace','remove','same-style'].includes(kind)?`<w:pStyle w:val="${kind==='same-style'?'Heading1':'Body'}"/>`:'';
 if(kind==='duplicate-pstyle')props='<w:pStyle w:val="Body"/><w:pStyle w:val="Heading1"/>';
 if(kind==='wrong-namespace')props='<w:pStyle val="Body"/>';
 if(kind==='property-revision')props='<w:pPrChange w:id="1" w:author="Earlier"><w:pPr/></w:pPrChange>';
 const ppr=kind==='misplaced-pstyle'?'<w:pPr><w:keepNext/><w:pStyle w:val="Body"/></w:pPr>':`<w:pPr>${props}<w:keepNext/><w:spacing w:after="160"/></w:pPr>`;
 p.set(p.mainPart(),p.text(p.mainPart()).replace(/(<w:p\b[^>]*>)/,`$1${ppr}${kind==='duplicate-properties'?'<w:pPr/>':''}`));
 addPart(p,'custom/style-sentinel.bin',new Uint8Array([4,7,9]),'application/octet-stream');
 if(kind==='protected'){addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');}
 return Document.open(p.toBytes());
}
export function directStyle(xml:string){const p=elements(parseXml(xml),'p',W)[0]!,pr=p.children.find(c=>c.namespaceURI===W&&c.localName==='pPr'),s=pr?.children.find(c=>c.namespaceURI===W&&c.localName==='pStyle');return s?attribute(s,'val',W):undefined;}
export const bindings:StepBinding[]=[
 {pattern:/^a native styled paragraph prepared for (assign|replace|remove|same-style|absent-removal)$/,run:async(c,op)=>{c.op=op;c.doc=await styleDocument(op);c.before=await(c.doc as Document).save();}},
 {pattern:/^the paragraph style is selected for (assign|replace|remove|same-style|absent-removal)$/,run:(c,op)=>{c.receipt=(c.doc as Document).paragraphs[0]!.setStyle(op==='remove'||op==='absent-removal'?null:'Heading1');}},
 {pattern:/^saved direct style and change receipt match (assign|replace|remove|same-style|absent-removal)$/,run:async(c,op)=>{const d=await Document.open(await(c.doc as Document).save()),expected=op==='remove'||op==='absent-removal'?undefined:'Heading1';expect(d.paragraphs[0]!.styleId).toBe(expected);expect(directStyle(new TextDecoder().decode(d.package.get('word/document.xml')!))).toBe(expected);expect(c.receipt).toEqual({changed:op==='same-style'||op==='absent-removal'?0:1});}},
 {pattern:/^paragraph text, other properties and unrelated package parts are preserved$/,run:async c=>{const d=c.doc as Document,before=await OpcPackage.open(c.before as Uint8Array),after=await OpcPackage.open(await d.save());expect(d.paragraphs[0]!.text).toBe('Styled π text');expect(after.names()).toEqual(before.names());for(const n of before.names())if(n!==before.mainPart())expect(after.get(n)).toEqual(before.get(n));expect(after.text(after.mainPart())).toContain('<w:keepNext/>');expect(after.text(after.mainPart())).toContain('<w:spacing w:after="160"/>');expect(after.text(after.mainPart())).toContain('<w:b/>');if(c.op==='same-style'||c.op==='absent-removal')expect(after.toBytes()).toEqual(c.before as Uint8Array);}},
 {pattern:/^an unsafe paragraph style input (\S+)$/,run:async(c,kind)=>{c.kind=kind;const d=await styleDocument(kind);c.doc=d;c.handle=d.paragraphs[0];if(kind==='stale')d.addParagraph('later');c.before=await d.save();}},
 {pattern:/^its paragraph style mutation is attempted$/,run:c=>{try{c.receipt=(c.handle as any).setStyle(c.kind==='unknown-style'?'Missing':c.kind==='character-style'?'Emphasis':c.kind==='invalid-argument'?17:'Heading1');}catch(e){c.error=e;}}},
 {pattern:/^style assignment refuses before changing package bytes$/,run:async c=>{expect(c.error).toBeDefined();expect((c.error as any).code).toMatch(/^docx-/);expect(await(c.doc as Document).save()).toEqual(Uint8Array.from(c.before as Uint8Array));}},
];
