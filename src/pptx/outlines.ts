import {OoxmlError} from '../errors.ts';
import {parseXml,elements,attribute,applyEdits,type XmlElement} from '../xml/index.ts';
import {retainedFormattingXml} from './formatting.ts';
export type OutlineCap='flat'|'rnd'|'sq';
export type OutlineCompound='sng'|'dbl'|'thickThin'|'thinThick'|'tri';
export type OutlineJoin={type:'round'|'bevel'}|{type:'miter';limit:number};
export type OutlineEnd={type:'none'|'triangle'|'stealth'|'diamond'|'oval'|'arrow';width:'sm'|'med'|'lg';length:'sm'|'med'|'lg'};
export type OutlineStyle={cap:OutlineCap;compound:OutlineCompound;join:OutlineJoin|null;headEnd:OutlineEnd|null;tailEnd:OutlineEnd|null};
export type OutlinePatch=Partial<OutlineStyle>;
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',N='http://www.w3.org/2000/xmlns/';
const caps=['flat','rnd','sq'],compounds=['sng','dbl','thickThin','thinThick','tri'],types=['none','triangle','stealth','diamond','oval','arrow'],sizes=['sm','med','lg'];
function fail(message:string):never{throw new OoxmlError('PPTX_OUTLINE_UNSUPPORTED',message);}
function plain(value:unknown,allowed:string[]):Record<string,unknown>{if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Outline requires plain data');const copy:Record<string,unknown>={};for(const key of Reflect.ownKeys(value)){if(typeof key!=='string'||!allowed.includes(key))fail('Unknown outline field');const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Outline accessors are unsupported');copy[key]=d.value;}return copy;}
function choice<T extends string>(value:unknown,values:readonly string[]):T{if(typeof value!=='string'||!values.includes(value))fail('Unsupported outline enum');return value as T;}
function join(value:unknown):OutlineJoin|null{if(value===null)return null;const j=plain(value,['type','limit']);if(j.type==='miter'){if(typeof j.limit!=='number'||!Number.isInteger(j.limit)||j.limit<0||j.limit>1000000)fail('Invalid miter limit');return {type:'miter',limit:j.limit};}if(Object.hasOwn(j,'limit'))fail('Only miter has a limit');return {type:choice<'round'|'bevel'>(j.type,['round','bevel'])};}
function end(value:unknown):OutlineEnd|null{if(value===null)return null;const e=plain(value,['type','width','length']);return {type:choice(e.type,types),width:choice(e.width,sizes),length:choice(e.length,sizes)};}
function request(value:unknown):OutlinePatch{const p=plain(value,['cap','compound','join','headEnd','tailEnd']),result:OutlinePatch={};for(const key of Object.keys(p)){if(key==='cap')result.cap=choice(p[key],caps);else if(key==='compound')result.compound=choice(p[key],compounds);else if(key==='join')result.join=join(p[key]);else result[key as 'headEnd'|'tailEnd']=end(p[key]);}return result;}
function attrs(n:XmlElement,allowed:string[],required:string[]=[]){const keys=Object.keys(n.attributes).filter(k=>n.attributeNamespaces[k]!==N);if(keys.some(k=>n.attributeNamespaces[k]!==''||!allowed.includes(k))||required.some(k=>!keys.includes(k)))fail('Unsupported line/decor attributes');}
function gaps(source:string,n:XmlElement){if(n.selfClosing)return;let cursor=n.openEnd;for(const c of n.children){if(source.slice(cursor,c.start).trim())fail('Outline lexical barrier');cursor=c.end;}if(source.slice(cursor,n.closeStart).trim())fail('Outline lexical barrier');}
function one(parent:XmlElement,name:string,ns=P):XmlElement{const rows=parent.children.filter(n=>n.namespaceURI===ns&&n.localName===name);if(rows.length!==1)fail('Expected unique '+name);return rows[0]!;}
function selected(source:string,id:number){
 if(!Number.isInteger(id)||id<1||id>2147483647)fail('Invalid outline ID');const doc=parseXml(source);if(doc.root.namespaceURI!==P||doc.root.localName!=='sld')fail('Expected slide root');const tree=one(one(doc.root,'cSld'),'spTree'),seen=new Set<number>();
 for(const n of elements(doc,'cNvPr',P)){const raw=attribute(n,'id'),value=raw&&/^\d+$/.test(raw)?Number(raw):NaN;if(!Number.isInteger(value)||value<1||value>2147483647||seen.has(value))fail('Duplicate/invalid outline identity');seen.add(value);}
 const matches=tree.children.filter(n=>n.namespaceURI===P&&['sp','cxnSp'].includes(n.localName)&&Number(attribute(one(one(n,n.localName==='sp'?'nvSpPr':'nvCxnSpPr'),'cNvPr'),'id'))===id);if(matches.length!==1)fail('Missing direct shape/connector outline');const line=one(one(matches[0]!,'spPr'),'ln',A);attrs(line,['w','cap','cmpd','algn']);gaps(source,line);
 if(attribute(line,'w')!==undefined&&(!/^\d+$/.test(attribute(line,'w')!)||Number(attribute(line,'w'))>20116800))fail('Invalid retained line width');if(attribute(line,'algn')!==undefined&&!['ctr','in'].includes(attribute(line,'algn')!))fail('Invalid retained line alignment');
 const fill=['noFill','solidFill','gradFill','pattFill'],dash=['prstDash','custDash'],joins=['round','bevel','miter'];let last=-1,hasFill=false,hasDash=false,hasJoin=false;const nodes:Record<string,XmlElement>={};
 for(const n of line.children){let rank=-1;if(fill.includes(n.localName)){if(hasFill)fail('Duplicate line fills');hasFill=true;rank=0;}else if(dash.includes(n.localName)){if(hasDash)fail('Duplicate line dashes');hasDash=true;rank=1;}else if(joins.includes(n.localName)){if(hasJoin)fail('Duplicate line joins');hasJoin=true;rank=2;nodes.join=n;}else if(n.localName==='headEnd'){rank=3;nodes.headEnd=n;}else if(n.localName==='tailEnd'){rank=4;nodes.tailEnd=n;}else if(n.localName==='extLst')rank=5;
  if(n.namespaceURI!==A||rank<0||rank<=last)fail('Unsupported line grammar/order');last=rank;
 }
 const j=nodes.join;let joinValue:OutlineJoin|null=null;if(j){gaps(source,j);if(j.children.length)fail('Decorated line join');attrs(j,j.localName==='miter'?['lim']:[],j.localName==='miter'?['lim']:[]);joinValue=join(j.localName==='miter'?{type:'miter',limit:/^\d+$/.test(attribute(j,'lim')??'')?Number(attribute(j,'lim')):NaN}:{type:j.localName});}
 function readEnd(n:XmlElement|undefined):OutlineEnd|null{if(!n)return null;attrs(n,['type','w','len']);gaps(source,n);if(n.children.length)fail('Decorated arrowhead');return end({type:attribute(n,'type')??'none',width:attribute(n,'w')??'med',length:attribute(n,'len')??'med'});}
 const value:OutlineStyle={cap:choice(attribute(line,'cap')??'flat',caps),compound:choice(attribute(line,'cmpd')??'sng',compounds),join:joinValue,headEnd:readEnd(nodes.headEnd),tailEnd:readEnd(nodes.tailEnd)};return {line,nodes,value};
}
export function getOutlineStyleXml(source:string,id:number):OutlineStyle{return selected(source,id).value;}
export function patchOutlineStyleXml(source:string,id:number,input:OutlinePatch):string{
 const patch=request(input),current=selected(source,id),target={...current.value,...patch};if(Object.keys(patch).every(k=>JSON.stringify(current.value[k as keyof OutlineStyle])===JSON.stringify(target[k as keyof OutlineStyle])))return source;
 let next=source;const lineAttrs:Record<string,string>={};if(target.cap!==current.value.cap)lineAttrs.cap=target.cap;if(target.compound!==current.value.compound)lineAttrs.cmpd=target.compound;
 if(Object.keys(lineAttrs).length)next=applyEdits(next,[{start:current.line.start,end:current.line.openEnd,value:retainedFormattingXml.openTag(next,current.line,lineAttrs)}]);
 for(const key of ['join','headEnd','tailEnd'] as const){if(!Object.hasOwn(patch,key)||JSON.stringify(target[key])===JSON.stringify(current.value[key]))continue;const {line,nodes}=selected(next,id),old=nodes[key],value=target[key];let content='';if(value){if(key==='join'){const j=value as OutlineJoin;content=`<a:${j.type} xmlns:a="${A}"${j.type==='miter'?` lim="${j.limit}"`:''}/>`;}else{const e=value as OutlineEnd;content=`<a:${key} xmlns:a="${A}" type="${e.type}" w="${e.width}" len="${e.length}"/>`;}}
  if(old&&value&&(key!=='join'||old.localName===(value as OutlineJoin).type)){
   const values:Record<string,string>=key==='join'?{lim:String((value as Extract<OutlineJoin,{type:'miter'}>).limit)}:{type:(value as OutlineEnd).type,w:(value as OutlineEnd).width,len:(value as OutlineEnd).length};
   if(key!=='join'||old.localName==='miter')next=applyEdits(next,[{start:old.start,end:old.openEnd,value:retainedFormattingXml.openTag(next,old,values)}]);
  }else if(old)next=applyEdits(next,[{start:old.start,end:old.end,value:content}]);else if(content){if(line.selfClosing)next=applyEdits(next,[{start:line.start,end:line.end,value:next.slice(line.start,line.openEnd).replace(/\/\s*>$/,()=>`>${content}</${line.name}>`)}]);else{const following=key==='join'?['headEnd','tailEnd','extLst']:key==='headEnd'?['tailEnd','extLst']:['extLst'],at=line.children.find(n=>following.includes(n.localName))?.start??line.closeStart;next=applyEdits(next,[{start:at,end:at,value:content}]);}}
 }
 return next;
}
