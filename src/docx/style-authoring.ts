import {OoxmlError} from '../errors.ts';
import {OpcPackage,addPart,addRelationship,getContentType,nextPartName} from '../opc/index.ts';
import {parseXml,applyEdits,attribute,escapeAttribute,type XmlElement} from '../xml/index.ts';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS='http://www.w3.org/2000/xmlns/';
const STYLE_REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const STYLE_TYPE='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const EFFECTS_REL='http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects';
export type AddParagraphStyleOptions={name:string;basedOn?:string;bold?:boolean;italic?:boolean};
export type ParagraphStyleDefinitionReceipt={styleId:string;partName:string};
const fail=(code:string,message:string):never=>{throw new OoxmlError(code,message);};
const word=(node:XmlElement,local:string)=>node.namespaceURI===W&&node.localName===local;

/** Copy plain data before package access; getters cannot mutate a package mid-operation. */
export function normalizeStyleOptions(styleId:string,options:AddParagraphStyleOptions):AddParagraphStyleOptions {
 for(const [label,value]of [['styleId',styleId],['name',options&&Object.getOwnPropertyDescriptor(options,'name')?.value]]){
  if(typeof value!=='string'||!value.trim()||value!==value.trim())fail('docx-style-argument',`Expected a nonempty ${label} without boundary whitespace`);
  escapeAttribute(value);
 }
 if(!options||typeof options!=='object'||Array.isArray(options)||![Object.prototype,null].includes(Object.getPrototypeOf(options)))fail('docx-style-argument','Style options must be plain data');
 const copy:Record<string,unknown>={};
 for(const key of Reflect.ownKeys(options)){
  if(typeof key!=='string'||!['name','basedOn','bold','italic'].includes(key))fail('docx-style-argument','Unknown paragraph style option');
  const descriptor=Object.getOwnPropertyDescriptor(options,key)!;
  if(!('value'in descriptor))fail('docx-style-argument','Style option accessors are unsupported');
  copy[key as string]=descriptor.value;
 }
 if(copy.basedOn!==undefined){if(typeof copy.basedOn!=='string'||!copy.basedOn.trim()||copy.basedOn!==copy.basedOn.trim())fail('docx-style-argument','Expected a nonempty base style ID');escapeAttribute(copy.basedOn as string);}
 for(const key of ['bold','italic'])if(copy[key]!==undefined&&typeof copy[key]!=='boolean')fail('docx-style-argument','Style bold/italic must be boolean');
 return copy as AddParagraphStyleOptions;
}

/** Appends one definition, or creates a styles part with matching relationship and MIME.
 * Does not rewrite existing definitions or interpret effective inherited formatting. */
export function authorParagraphStyle(pkg:OpcPackage,owner:string,styleId:string,options:AddParagraphStyleOptions):ParagraphStyleDefinitionReceipt {
 const rels=pkg.relationships(owner);
 if(rels.some(r=>r.type===EFFECTS_REL))fail('docx-style-unsupported','Dual stylesWithEffects registries require reconciliation');
 const links=rels.filter(r=>r.type===STYLE_REL);
 if(links.length>1||links.some(r=>r.external))fail('docx-style-unsupported','Styles relationship must be unique and internal');
 let partName=links[0]?.resolved;
 let xml:string;
 if(partName){
  if(getContentType(pkg,partName)!==STYLE_TYPE)fail('docx-style-part-invalid','Unexpected styles content type');
  xml=pkg.text(partName);
 }else{
  partName=pkg.names().some(n=>n.toLowerCase()==='word/styles.xml')?nextPartName(pkg,'word/styles%d.xml'):'word/styles.xml';
  xml=`<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${W}"/>`;
 }
 const root=parseXml(xml).root;
 if(!word(root,'styles'))fail('docx-style-part-invalid','Invalid styles root');
 const styles=new Map<string,XmlElement>();let phase=0;
 for(const child of root.children){
  if(word(child,'docDefaults')&&phase===0){phase=1;continue;}
  if(word(child,'latentStyles')&&phase<2){phase=2;continue;}
  if(!word(child,'style'))fail('docx-style-unsupported','Unsupported styles registry child or order');
  phase=3;
  const id=attribute(child,'styleId',W),type=attribute(child,'type',W);
  if(!id?.trim()||!['paragraph','character','table','numbering'].includes(type??''))fail('docx-style-part-invalid','Every existing style needs a qualified ID and supported type');
  if(styles.has(id!))fail('docx-style-ambiguous','Duplicate registry style ID');
  styles.set(id!,child);
 }
 // Appending must not silently legitimise a malformed registry with raw text.
 let cursor=root.openEnd;
 if(!root.selfClosing){for(const child of root.children){if(!/^[\t\r\n ]*$/.test(xml.slice(cursor,child.start)))fail('docx-style-unsupported','Registry lexical barriers are unsupported');cursor=child.end;}if(!/^[\t\r\n ]*$/.test(xml.slice(cursor,root.closeStart)))fail('docx-style-unsupported','Registry lexical barriers are unsupported');}
 if(styles.has(styleId))fail('docx-style-duplicate','Style ID already exists');
 const seen=new Set<string>([styleId]);let base=options.basedOn;
 while(base!==undefined){
  if(seen.has(base))fail('docx-style-base-invalid','Cyclic paragraph style base chain');seen.add(base);
  const node=styles.get(base);
  if(!node||attribute(node,'type',W)!=='paragraph')fail('docx-style-base-invalid','Base chain must resolve to paragraph styles');
  const bases=node!.children.filter(c=>c.localName==='basedOn');
  if(bases.length>1)fail('docx-style-base-invalid','Ambiguous base style');
  const link=bases[0];base=undefined;
  if(link){
   if(!word(link,'basedOn')||link.children.length||!link.selfClosing&&xml.slice(link.openEnd,link.closeStart).trim())fail('docx-style-base-invalid','Malformed basedOn property');
   for(const key of Object.keys(link.attributes))if(link.attributeNamespaces[key]!==XMLNS&&!(key.split(':').at(-1)==='val'&&link.attributeNamespaces[key]===W))fail('docx-style-base-invalid','Malformed basedOn attribute');
   base=attribute(link,'val',W);
   if(!base?.trim())fail('docx-style-base-invalid','Missing base style ID');
  }
 }
 const rpr=(['bold','italic'] as const).filter(k=>options[k]!==undefined).map(k=>`<w:${k==='bold'?'b':'i'} w:val="${options[k]?'1':'0'}"/>`).join('');
 const definition=`<w:style xmlns:w="${W}" w:type="paragraph" w:customStyle="1" w:styleId="${escapeAttribute(styleId)}"><w:name w:val="${escapeAttribute(options.name)}"/>${options.basedOn===undefined?'':`<w:basedOn w:val="${escapeAttribute(options.basedOn)}"/>`}${rpr?`<w:rPr>${rpr}</w:rPr>`:''}</w:style>`;
 const next=root.selfClosing?applyEdits(xml,[{start:root.start,end:root.end,value:xml.slice(root.start,root.openEnd).replace(/\/\s*>$/,()=>`>${definition}</${root.name}>`)}]):applyEdits(xml,[{start:root.closeStart,end:root.closeStart,value:definition}]);
 const target=partName;
 pkg.transaction(()=>{
  if(links.length)pkg.set(target,next);
  else{addPart(pkg,target,next,STYLE_TYPE);addRelationship(pkg,owner,STYLE_REL,'/'+target);}
  pkg.toBytes();
 });
 return {styleId,partName:target};
}
