import {OoxmlError} from '../errors.ts';
import {OpcPackage,getContentType} from '../opc/index.ts';
import {parseXml,applyEdits,attribute} from '../xml/index.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',XMLNS='http://www.w3.org/2000/xmlns/';
const fail=(message:string):never=>{throw new OoxmlError('PPTX_ORDER_UNSUPPORTED',message);};
export function permutation(value:readonly number[],count:number):number[]{
 if(!Array.isArray(value)||Object.getPrototypeOf(value)!==Array.prototype||value.length!==count)fail('Order must be an exact array permutation');
 if(Reflect.ownKeys(value).some(k=>k!=='length'&&(typeof k!=='string'||!/^\d+$/.test(k)||String(Number(k))!==k||Number(k)>=count)))fail('Unknown order property');
 const order:number[]=[],seen=new Set<number>();for(let i=0;i<count;i++){const d=Object.getOwnPropertyDescriptor(value,String(i));if(!d||!('value'in d)||!Number.isInteger(d.value)||d.value<0||d.value>=count||seen.has(d.value))fail('Order must contain every current index exactly once as plain data');seen.add(d!.value);order.push(d!.value);}return order;
}
/** Replace only complete list entries at their existing slots; keep surrounding XML. */
export function reorderSlideList(pkg:OpcPackage,main:string,parts:readonly string[],order:readonly number[]):string{
 const xml=pkg.text(main),root=parseXml(xml).root;
 if(root.namespaceURI!==P||root.localName!=='presentation')fail('Invalid presentation root');
 if(root.children.some(n=>n.namespaceURI!==P||!['sldMasterIdLst','notesMasterIdLst','handoutMasterIdLst','sldIdLst','sldSz','notesSz','embeddedFontLst','kinsoku','defaultTextStyle'].includes(n.localName)))fail('Custom shows, extensions, protection or unknown presentation metadata require explicit reconciliation');
 const lists=root.children.filter(n=>n.localName==='sldIdLst');if(lists.length!==1)fail('Expected one slide list');const list=lists[0]!;
 if(Object.keys(list.attributes).some(k=>list.attributeNamespaces[k]!==XMLNS))fail('Slide list attributes are unsupported');
 let cursor=list.openEnd;const ids=new Set<number>(),rids=new Set<string>(),targets=new Set<string>();const rels=pkg.relationships(main);
 if(list.children.length!==parts.length)fail('Slide count differs from handles');
 for(const [i,node]of list.children.entries()){
  if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,node.start)))fail('Slide-list lexical barriers are unsupported');cursor=node.end;
  if(node.namespaceURI!==P||node.localName!=='sldId'||node.children.length||!node.selfClosing&&xml.slice(node.openEnd,node.closeStart).trim())fail('Unsupported slide-list entry');
  for(const key of Object.keys(node.attributes))if(node.attributeNamespaces[key]!==XMLNS&&!(key==='id'||key.split(':').at(-1)==='id'&&node.attributeNamespaces[key]===R))fail('Unsupported slide-ID attribute');
  const raw=node.attributes.id,id=raw&&/^\d+$/.test(raw)?Number(raw):NaN,rid=attribute(node,'id',R);if(!Number.isInteger(id)||id<256||id>2147483647||ids.has(id)||!rid||rids.has(rid))fail('Invalid/duplicate slide ID or relationship ID');
  ids.add(id);rids.add(rid!);const link=rels.find(r=>r.id===rid);
  if(!link||link.type!==R+'/slide'||link.external||!link.resolved||link.resolved!==parts[i]||targets.has(link.resolved))fail('Ambiguous slide relationship or target');
  if(getContentType(pkg,link!.resolved!)!=='application/vnd.openxmlformats-officedocument.presentationml.slide+xml')fail('Wrong slide content type');targets.add(link!.resolved!);
 }
 if(!list.selfClosing&&!/^[ \t\r\n]*$/.test(xml.slice(cursor,list.closeStart)))fail('Slide-list lexical barriers are unsupported');
 if(order.every((n,i)=>n===i))return xml;
 return applyEdits(xml,list.children.map((node,i)=>({start:node.start,end:node.end,value:xml.slice(list.children[order[i]!]!.start,list.children[order[i]!]!.end)})));
}
