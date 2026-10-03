import {OoxmlError} from '../errors.ts';
import {parseXml,elements,attribute,escapeAttribute,applyEdits,type XmlElement} from '../xml/index.ts';
import {shapeAppendSite} from './text-box.ts';
import type {PictureGeometry} from './picture-add.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main';
export type ShapeGroupOptions={name?:string};
export type ShapeGroupReceipt={shapeId:number;partName:string;childIds:number[];geometry:PictureGeometry};
function fail(message:string):never{throw new OoxmlError('PPTX_GROUP_UNSUPPORTED',message);}
function plain(value:unknown,keys:string[]):Record<string,unknown>{if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Group inputs require plain data');const result:Record<string,unknown>={};for(const k of Reflect.ownKeys(value)){if(typeof k!=='string'||!keys.includes(k))fail('Unknown group field');const d=Object.getOwnPropertyDescriptor(value,k)!;if(!('value'in d))fail('Group accessors are unsupported');result[k]=d.value;}return result;}
function identity(node:XmlElement):number {const nv=node.children.find(n=>n.namespaceURI===P&&['nvSpPr','nvPicPr'].includes(n.localName)),id=nv?.children.find(n=>n.namespaceURI===P&&n.localName==='cNvPr');return Number(id&&attribute(id,'id'));}
export function groupShapesXml(source:string,ids:number[],geometry:PictureGeometry,options:ShapeGroupOptions={}){
 if(!Array.isArray(ids)||ids.length<2||ids.length>100)fail('Group requires 2–100 unique shape IDs');const selectedIds:number[]=[];
 for(let index=0;index<ids.length;index++){const d=Object.getOwnPropertyDescriptor(ids,String(index));if(!d||!('value'in d)||!Number.isInteger(d.value)||d.value<1||d.value>2147483647||selectedIds.includes(d.value))fail('Invalid or duplicate group selection');selectedIds.push(d.value);}
 const g=plain(geometry,['x','y','width','height']),o=plain(options,['name']);for(const k of ['x','y','width','height']){const v=g[k];if(typeof v!=='number'||!Number.isInteger(v)||v<(k==='x'||k==='y'?-2147483648:1)||v>2147483647)fail('Group geometry requires bounded integers');}
 if(o.name!==undefined){if(typeof o.name!=='string'||!o.name.trim())fail('Group name must be nonempty text');try{escapeAttribute(o.name);}catch{fail('Invalid group name');}}
 let id:number;try{id=shapeAppendSite(source).shapeId;}catch(error){if(error instanceof OoxmlError&&error.code==='PPTX_TEXT_BOX_UNSUPPORTED')fail(error.message);throw error;}
 const doc=parseXml(source),tree=elements(doc,'spTree',P)[0]!,children=tree.children.filter(n=>n.namespaceURI===P&&['sp','pic'].includes(n.localName)&&selectedIds.includes(identity(n)));
 if(children.length!==selectedIds.length)fail('Group requires direct shape/picture children');const indices=children.map(n=>tree.children.indexOf(n));if(indices.some((n,index)=>n!==indices[0]!+index))fail('Group selection must be contiguous in source order');
 for(const child of children){if(elements(child,'ph',P).length)fail('Grouping placeholders is unsupported');for(const lock of [...elements(child,'spLocks',A),...elements(child,'picLocks',A)])if(!['0','false',undefined].includes(attribute(lock,'noGrp')))fail('Shape grouping is locked');}
 for(const n of [...elements(doc,'stCxn',A),...elements(doc,'endCxn',A)])if(selectedIds.includes(Number(attribute(n,'id'))))fail('Grouping attached connector endpoints is unsupported');
 // Avoid introducing a binding that changes a prefix inherited by existing children.
 const lexicalPrefixes=new Set(doc.elements.flatMap(n=>[n.name,...Object.keys(n.attributes)].filter(k=>k.includes(':')).flatMap(k=>k.startsWith('xmlns:')?[k.slice(6)]:[k.split(':')[0]!])));
 function prefix(base:string){let n=0,key=base;while(lexicalPrefixes.has(key)||doc.root.attributes['xmlns:'+key]!==undefined)key=base+ ++n;lexicalPrefixes.add(key);return key;}
 const p=prefix('group'),a=prefix('draw'),first=children[0]!,last=children.at(-1)!,body=source.slice(first.start,last.end),name=o.name??`Group ${id}`;
 const wrapper=`<${p}:grpSp xmlns:${p}="${P}" xmlns:${a}="${A}"><${p}:nvGrpSpPr><${p}:cNvPr id="${id}" name="${escapeAttribute(name as string)}"/><${p}:cNvGrpSpPr/><${p}:nvPr/></${p}:nvGrpSpPr><${p}:grpSpPr><${a}:xfrm><${a}:off x="${g.x}" y="${g.y}"/><${a}:ext cx="${g.width}" cy="${g.height}"/><${a}:chOff x="${g.x}" y="${g.y}"/><${a}:chExt cx="${g.width}" cy="${g.height}"/></${a}:xfrm></${p}:grpSpPr>${body}</${p}:grpSp>`;
 return {xml:applyEdits(source,[{start:first.start,end:last.end,value:wrapper}]),shapeId:id,childIds:children.map(identity),geometry:g as PictureGeometry};
}
