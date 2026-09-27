import {attribute,type XmlElement} from '../xml/index.ts';
import {runPropertySetSupported} from './revision-properties.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/',XML='http://www.w3.org/XML/1998/namespace';
export const MOVE_NAMES=new Set(['moveFrom','moveTo','moveFromRangeStart','moveFromRangeEnd','moveToRangeStart','moveToRangeEnd']);
type Move={element:XmlElement;kind:'move-from'|'move-to';text:string};
export type MoveScan={moves:Move[];markers:XmlElement[];identifiers:XmlElement[];problems:string[]};
const word=(e:XmlElement|undefined,name:string)=>e?.namespaceURI===W&&e.localName===name;
const id=(e:XmlElement)=>attribute(e,'id',W);
const validId=(v:string|undefined)=>v!==undefined&&/^\d+$/.test(v)&&Number.isSafeInteger(Number(v));
function attrs(e:XmlElement,allowed:string[]):boolean{return Object.keys(e.attributes).every(k=>e.attributeNamespaces[k]===XMLNS||(e.attributeNamespaces[k]===W&&allowed.includes(k.split(':').pop()!)));}
function whitespace(e:XmlElement,xml:string):boolean{if(e.selfClosing)return true;let at=e.openEnd;for(const c of e.children){if(!/^[ \t\r\n]*$/.test(xml.slice(at,c.start)))return false;at=c.end;}return /^[ \t\r\n]*$/.test(xml.slice(at,e.closeStart));}
function canonical(e:XmlElement):unknown{return [e.namespaceURI,e.localName,Object.keys(e.attributes).filter(k=>e.attributeNamespaces[k]!==XMLNS).map(k=>[e.attributeNamespaces[k],k.split(':').pop(),e.attributes[k]]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b))),e.children.map(canonical)];}
// Deliberately bounded UTC lexical profile; no timezone/date normalisation on write.
function validDate(value:string|undefined):boolean {
 if(!value||!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(value)||!Number.isFinite(Date.parse(value)))return false;
 const normal=new Date(value).toISOString(),expected=value.includes('.')?value.replace(/\.(\d+)Z$/,(_,f:string)=>'.'+f.padEnd(3,'0')+'Z'):value.replace('Z','.000Z');return normal===expected;
}
function plainMove(wrapper:XmlElement,xml:string):{signature:string;text:string}|undefined{
 const date=attribute(wrapper,'date',W);if(date!==undefined&&!validDate(date))return;
 if(!wrapper.children.length||!whitespace(wrapper,xml)||!attrs(wrapper,['id','author','date'])||attribute(wrapper,'author',W)===undefined)return;
 const values:unknown[]=[];let text='';
 for(const run of wrapper.children){if(!word(run,'r')||!attrs(run,[])||!whitespace(run,xml))return;let pr:XmlElement|undefined;const texts:unknown[]=[];
  for(const child of run.children){if(word(child,'rPr')){if(pr||child!==run.children[0]||!runPropertySetSupported(child,xml))return;pr=child;continue;}
   if(!word(child,'t')||child.children.length||(!child.selfClosing&&/<!--|<!\[CDATA\[|<\?/.test(xml.slice(child.openEnd,child.closeStart))))return;
   if(Object.keys(child.attributes).some(k=>child.attributeNamespaces[k]!==XMLNS&&(child.attributeNamespaces[k]!==XML||k.split(':').pop()!=='space'||!['default','preserve'].includes(child.attributes[k]!))))return;
   texts.push([child.text,attribute(child,'space',XML)??null]);text+=child.text;
  }
  if(!texts.length)return;values.push([pr?pr.children.map(canonical):[],texts]);
 }
 if(!text.length)return;
 return {signature:JSON.stringify(values),text};
}
/** Narrow same-story range pairs; all move findings block mutation of the story. */
export function scanRunMoves(elements:XmlElement[],xml:string):MoveScan{
 const nodes=elements.filter(e=>e.namespaceURI===W&&MOVE_NAMES.has(e.localName)),result:MoveScan={moves:[],markers:[],identifiers:[],problems:[]};if(!nodes.length)return result;
 const ranges=new Map<string,{from?:{wrapper:XmlElement;signature:string;text:string};to?:{wrapper:XmlElement;signature:string;text:string}}>(),owned=new Set<XmlElement>(),ids=new Set<string>();
 const problem=(s:string)=>result.problems.push(s);
 for(const start of nodes.filter(n=>n.localName.endsWith('RangeStart'))){
  const from=start.localName==='moveFromRangeStart',parent=start.parent,kind=from?'moveFrom':'moveTo',name=attribute(start,'name',W),index=parent?.children.indexOf(start)??-1,wrapper=parent?.children[index+1],end=parent?.children[index+2];
  if(!word(parent,'p')||index<0||!word(wrapper,kind)||!word(end,kind+'RangeEnd')){problem('move range must be a direct paragraph start, single wrapper and matching end triple');continue;}
  let nested=false;for(let n=parent!.parent;n;n=n.parent)if(n.namespaceURI===W&&(MOVE_NAMES.has(n.localName)||['ins','del'].includes(n.localName)||n.localName.endsWith('Change')))nested=true;
  if(nested||!attrs(start,['id','name','author','date'])||attribute(start,'author',W)===undefined||!validDate(attribute(start,'date',W))||!name||!name.trim()||!attrs(end!,['id'])||start.children.length||end!.children.length||!whitespace(start,xml)||!whitespace(end!,xml)||!/^[ \t\r\n]*$/.test(xml.slice(start.end,wrapper!.start))||!/^[ \t\r\n]*$/.test(xml.slice(wrapper!.end,end!.start))){problem('move range metadata, nesting or lexical content is unsupported');continue;}
  if(!validId(id(start))||!validId(id(end!))||Number(id(start))!==Number(id(end!))||!validId(id(wrapper!))){problem('move range/wrapper IDs must be matching nonnegative safe decimals');continue;}
  const data=plainMove(wrapper!,xml);if(!data){problem('moved runs must contain matching bounded direct properties and ordinary text');continue;}
  const group=ranges.get(name)??{},side=from?'from':'to';if(group[side]){problem('move names must have exactly one source and one destination range per story');continue;}
  group[side]={wrapper:wrapper!,...data};ranges.set(name,group);
  for(const node of [start,wrapper!]){const key=String(Number(id(node)));if(ids.has(key))problem('move range and wrapper identifiers must be unique within a story');ids.add(key);result.identifiers.push(node);}
  for(const node of [start,wrapper!,end!]){if(owned.has(node))problem('move ranges must not overlap');owned.add(node);}result.markers.push(start,end!);
 }
 if(nodes.some(n=>!owned.has(n)))problem('unpaired, nested or non-run move markup is unsupported');
 for(const group of ranges.values()){if(!group.from||!group.to){problem('move source and destination with the same name must be in one story');continue;}if(group.from.signature!==group.to.signature){problem('move source/destination run text and direct properties must match');continue;}result.moves.push({element:group.from.wrapper,kind:'move-from',text:group.from.text},{element:group.to.wrapper,kind:'move-to',text:group.to.text});}
 if(result.problems.length){result.moves=[];result.markers=[];result.identifiers=[];}
 return result;
}
