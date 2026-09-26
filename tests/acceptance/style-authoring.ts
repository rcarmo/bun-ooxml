import {expect} from 'bun:test';
import {Document, type AddParagraphStyleOptions} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship,getContentType} from '../../src/opc/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {styleDocument,W,STYLE_TYPE,STYLE_REL} from './paragraph-style.ts';
export async function authoringDocument(kind:string){
 let d:Document;
 if(kind==='absent'||kind==='collision'){d=Document.create();d.addParagraph('Authored text');}
 else d=await styleDocument(kind);
 const p=await OpcPackage.open(await d.save());
 if(kind==='empty')p.set('word/styles.xml',`<w:styles xmlns:w="${W}"/>`);
 if(kind==='aliased')p.set('word/styles.xml',p.text('word/styles.xml').replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:'));
 if(kind==='default-namespace')p.set('word/styles.xml',p.text('word/styles.xml').replaceAll('<w:','<').replaceAll('</w:','</').replace('<styles ',`<styles xmlns="${W}" `));
 if(kind==='collision')addPart(p,'word/styles.xml','opaque orphan','application/octet-stream');
 if(['cyclic-base','broken-base-chain','malformed-base','duplicate-base','wrong-id-namespace'].includes(kind)){
  let xml=p.text('word/styles.xml');
  if(kind==='duplicate-base')xml=xml.replace('</w:styles>','<w:style w:type="paragraph" w:styleId="Heading1"/></w:styles>');
  else if(kind==='wrong-id-namespace')xml=xml.replace('w:styleId="Body"','styleId="Body"');
  else xml=xml.replace('<w:name w:val="Heading One"/>',`<w:name w:val="Heading One"/><w:basedOn ${kind==='malformed-base'?'val':'w:val'}="${kind==='cyclic-base'?'Heading1':'Missing'}"/>`);
  p.set('word/styles.xml',xml);
 }
 if(kind==='wrong-root')p.set('word/styles.xml','<wrong/>');
 if(kind==='styles-with-effects'){addPart(p,'word/stylesWithEffects.xml',`<w:styles xmlns:w="${W}"/>`,STYLE_TYPE);addRelationship(p,p.mainPart(),'http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects','stylesWithEffects.xml');}
 addPart(p,'custom/authoring-sentinel.bin',new Uint8Array([5,9,13]),'application/octet-stream');
 return Document.open(p.toBytes());
}
export function styleOptions(kind:string):AddParagraphStyleOptions {
 return {name:'Custom & π',...(kind==='based'?{basedOn:'Heading1'}:{}),...(kind==='flags'?{bold:true,italic:false}:{})};
}
export const bindings:StepBinding[]=[
 {pattern:/^a native style authoring document with (\S+)$/,run:async(c,kind)=>{c.doc=await authoringDocument(kind!);c.before=await(c.doc as Document).save();}},
 {pattern:/^a named paragraph style is authored with (plain|based|flags)$/,run:(c,kind)=>{c.receipt=(c.doc as Document).addParagraphStyle('Custom',styleOptions(kind!));}},
 {pattern:/^its saved definition and relationship are correct for (plain|based|flags)$/,run:async(c,kind)=>{
  const p=await OpcPackage.open(await(c.doc as Document).save()),links=p.relationships(p.mainPart()).filter(r=>r.type===STYLE_REL);expect(links.length).toBe(1);expect(links[0]!.external).toBe(false);const part=links[0]!.resolved!;expect(getContentType(p,part)).toBe(STYLE_TYPE);expect(c.receipt).toEqual({styleId:'Custom',partName:part});
  const root=parseXml(p.text(part)),defs=root.root.children.filter(n=>n.namespaceURI===W&&n.localName==='style'&&attribute(n,'styleId',W)==='Custom');expect(defs.length).toBe(1);const def=defs[0]!;expect(attribute(def,'type',W)).toBe('paragraph');expect(attribute(def,'customStyle',W)).toBe('1');expect(attribute(elements(def,'name',W)[0]!,'val',W)).toBe('Custom & π');const bases=elements(def,'basedOn',W);expect(bases.length).toBe(kind==='based'?1:0);if(kind==='based')expect(attribute(bases[0]!,'val',W)).toBe('Heading1');for(const name of ['b','i']){const nodes=elements(def,name,W);expect(nodes.length).toBe(kind==='flags'?1:0);if(kind==='flags')expect(attribute(nodes[0]!,'val',W)).toBe(name==='b'?'1':'0');}
 }},
 {pattern:/^pre-existing style definitions and unrelated package bytes are preserved$/,run:async c=>{
  const before=await OpcPackage.open(c.before as Uint8Array),after=await OpcPackage.open(await(c.doc as Document).save());const old=before.relationships(before.mainPart()).find(r=>r.type===STYLE_REL)?.resolved;const allowed=old?[old]:['[Content_Types].xml','word/_rels/document.xml.rels',(c.receipt as any).partName];expect(after.names().filter(n=>!before.names().includes(n)).sort()).toEqual(allowed.filter(n=>!before.names().includes(n)).sort());for(const n of before.names())if(!allowed.includes(n))expect(after.get(n)).toEqual(before.get(n));if(old){const x=before.text(old),y=after.text(old);for(const n of parseXml(x).root.children)expect(y).toContain(x.slice(n.start,n.end));}
 }},
 {pattern:/^the authored style can be selected and reopened$/,run:async c=>{const d=c.doc as Document,text=d.paragraphs[0]!.text;d.paragraphs[0]!.setStyle('Custom');const q=await Document.open(await d.save());expect(q.paragraphs[0]!.styleId).toBe('Custom');expect(q.paragraphs[0]!.text).toBe(text);}},
 {pattern:/^an unsafe native style authoring input (\S+)$/,run:async(c,kind)=>{const d=await authoringDocument(kind!);c.doc=d;c.kind=kind;c.handle=d.paragraphs[0];c.version=d.currentVersion();if(kind==='stale-document'){const p=await OpcPackage.open(await d.save());d.package.setPart('word/document.xml',new TextEncoder().encode(p.text(p.mainPart()).replace('Styled','External')));}c.before=await d.save();}},
 {pattern:/^its paragraph style creation is attempted$/,run:c=>{const k=c.kind;try{c.receipt=(c.doc as Document).addParagraphStyle(k==='duplicate-id'?'Heading1':k==='character-id'?'Emphasis':'Custom',k==='invalid-argument'?{name:'Bad',bold:'yes'} as any:{name:'New',basedOn:k==='missing-base'?'Missing':k==='character-base'?'Emphasis':'Heading1'});}catch(e){c.error=e;}}},
 {pattern:/^style authoring refuses without changing archive bytes or handle state$/,run:async c=>{expect(c.error).toBeDefined();expect((c.error as any).code).toMatch(/^docx-/);const d=c.doc as Document;expect(await d.save()).toEqual(Uint8Array.from(c.before as Uint8Array));expect(d.currentVersion()).toBe(c.version as number);expect(d.paragraphs[0]).toBe(c.handle as Document['paragraphs'][number]);}},
];
