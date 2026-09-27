import {OoxmlError} from '../errors.ts';
import {parseXml,applyEdits,escapeText,attribute,type XmlElement} from '../xml/index.ts';
import {OpcPackage} from './package.ts';
import {addPart,addRelationship} from './graph.ts';
import {getContentType} from './content-types.ts';
const CP='http://schemas.openxmlformats.org/package/2006/metadata/core-properties',DC='http://purl.org/dc/elements/1.1/',DCT='http://purl.org/dc/terms/',XSI='http://www.w3.org/2001/XMLSchema-instance',XMLNS='http://www.w3.org/2000/xmlns/';
const REL='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',MIME='application/vnd.openxmlformats-package.core-properties+xml',PATH='docProps/core.xml';
const FIELDS={title:DC,creator:DC,subject:DC,description:DC,identifier:DC,language:DC,keywords:CP,category:CP,contentStatus:CP,lastModifiedBy:CP,revision:CP,version:CP,created:DCT,modified:DCT,lastPrinted:CP} as const;
type Field=keyof typeof FIELDS;
export type CoreProperties={ [K in Field]:string|null };
export type CorePropertiesPatch=Partial<CoreProperties>;
const keys=Object.keys(FIELDS) as Field[];
const dates=new Set<Field>(['created','modified','lastPrinted']);
function fail(message:string):never{throw new OoxmlError('opc-core-properties-unsupported',message);}
function value(key:Field,input:unknown):string{
 if(typeof input!=='string'||input.length>4096)fail('Core values must be strings of at most 4096 UTF-16 units');escapeText(input);
 if(dates.has(key)){
  if(!/^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/.test(input)||input.startsWith('0000'))fail('Timestamp must be a full UTC second value');
  const parsed=new Date(input);if(!Number.isFinite(parsed.getTime())||parsed.toISOString()!==input.slice(0,-1)+'.000Z')fail('Invalid calendar timestamp');
 }
 return input;
}
function patch(input:CorePropertiesPatch):CorePropertiesPatch{
 if(!input||typeof input!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(input)))fail('Expected plain core-property patch');const result:CorePropertiesPatch={};
 for(const key of Reflect.ownKeys(input)){if(typeof key!=='string'||!Object.hasOwn(FIELDS,key))fail('Unknown core property');const d=Object.getOwnPropertyDescriptor(input,key)!;if(!('value'in d))fail('Core properties cannot be accessors');if(d.value===undefined)continue;Object.assign(result,{[key]:d.value===null?null:value(key as Field,d.value)});}
 if(!Object.keys(result).length)fail('Expected at least one defined core property');return result;
}
function whitespace(xml:string,node:XmlElement){if(node.selfClosing)return;let pos=node.openEnd;for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(pos,child.start)))fail('Mixed or lexical core-property content');pos=child.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(pos,node.closeStart)))fail('Mixed or lexical core-property content');}
function namespace(node:XmlElement,prefix:string):string|undefined{for(let n:XmlElement|undefined=node;n;n=n.parent){const v=n.attributes[prefix?'xmlns:'+prefix:'xmlns'];if(v!==undefined)return v;}return undefined;}
function inspect(pkg:OpcPackage){
 const links=pkg.relationships('').filter(r=>r.type===REL),parts=pkg.names().filter(n=>n!=='[Content_Types].xml'&&!n.endsWith('.rels')&&getContentType(pkg,n)===MIME);
 if(links.length>1||parts.length>1)fail('Ambiguous core-property parts or relationships');
 const link=links[0];if(!link){if(parts.length)fail('Orphan core-property part');return undefined;}
 if(link.external||!link.resolved||!pkg.get(link.resolved)||getContentType(pkg,link.resolved)!==MIME||parts[0]!==link.resolved)fail('Invalid core-property relationship or content type');
 const path=link.resolved,xml=pkg.text(path),root=parseXml(xml).root;
 if(root.namespaceURI!==CP||root.localName!=='coreProperties')fail('Invalid core-property root');
 for(const key of Object.keys(root.attributes))if(root.attributeNamespaces[key]!==XMLNS)fail('Core-property root attributes are unsupported');whitespace(xml,root);
 const nodes=new Map<Field,XmlElement>(),values=Object.fromEntries(keys.map(k=>[k,null])) as CoreProperties;
 for(const node of root.children){const key=node.localName as Field;if(!Object.hasOwn(FIELDS,key)||node.namespaceURI!==FIELDS[key]||nodes.has(key)||node.children.length)fail('Unknown, duplicate or complex core property');
  if(!node.selfClosing&&xml.slice(node.openEnd,node.closeStart).includes('<'))fail('Lexical core-property value cannot be rewritten safely');
  const typed=key==='created'||key==='modified';
  for(const name of Object.keys(node.attributes))if(node.attributeNamespaces[name]!==XMLNS&&!(typed&&node.attributeNamespaces[name]===XSI&&name.split(':').at(-1)==='type'))fail('Unsupported core-property attributes');
  if(typed){const qname=attribute(node,'type',XSI),bits=qname?.split(':');if(!bits||bits.length!==2||bits[1]!=='W3CDTF'||namespace(node,bits[0]!)!==DCT)fail('Created/modified require a qualified dcterms:W3CDTF type');}
  values[key]=value(key,node.text);nodes.set(key,node);
 }
 return {path,xml,root,nodes,values};
}
export function readCoreProperties(pkg:OpcPackage):CoreProperties{return inspect(pkg)?.values??Object.fromEntries(keys.map(k=>[k,null])) as CoreProperties;}
function fragment(key:Field,text:string):string{
 const prefix=FIELDS[key]===DC?'dc':FIELDS[key]===DCT?'dcterms':'cp';const typed=key==='created'||key==='modified';
 return `<${prefix}:${key} xmlns:${prefix}="${FIELDS[key]}"${typed?` xmlns:xsi="${XSI}" xsi:type="dcterms:W3CDTF"`:''}>${escapeText(text)}</${prefix}:${key}>`;
}
export function patchCoreProperties(pkg:OpcPackage,input:CorePropertiesPatch):{changed:number}{
 const values=patch(input),current=inspect(pkg),edits:Array<{start:number;end:number;value:string}>=[],additions:string[]=[];let changed=0;
 for(const key of keys){const next=values[key];if(next===undefined||next===(current?.values[key]??null))continue;changed++;const node=current?.nodes.get(key),rendered=next===null?'':fragment(key,next);if(node)edits.push({start:node.start,end:node.end,value:rendered});else additions.push(rendered);}
 if(!changed)return {changed:0};
 let nextXml:string;
 if(!current){if(pkg.names().some(n=>n.toLowerCase()===PATH.toLowerCase()))fail('Default core-property path is occupied');nextXml=`<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="${CP}">${additions.join('')}</cp:coreProperties>`;}
 else{const{root,xml}=current;if(additions.length){if(root.selfClosing)edits.push({start:root.start,end:root.end,value:xml.slice(root.start,root.end).replace(/\/>$/,'>')+additions.join('')+`</${root.name}>`});else edits.push({start:root.closeStart,end:root.closeStart,value:additions.join('')});}nextXml=applyEdits(xml,edits);}
 // Creation and update are one transaction, including relationship/type metadata.
 pkg.transaction(()=>{if(current){
  const original=pkg.get(current.path)!;
  // OPC's string setter retains UTF-16 but does not restore a UTF-8 BOM.
  const bom=original[0]===239&&original[1]===187&&original[2]===191?'\ufeff':'';
  pkg.set(current.path,bom+nextXml);
 }else{addPart(pkg,PATH,nextXml,MIME);addRelationship(pkg,'',REL,PATH);}const observed=readCoreProperties(pkg);for(const key of keys){const expected=values[key]===undefined?(current?.values[key]??null):values[key];if(observed[key]!==expected)fail('Core-property readback mismatch');}pkg.toBytes();});
 return {changed};
}
