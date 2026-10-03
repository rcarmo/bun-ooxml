import {OoxmlError} from '../errors.ts';
import {parseXml,elements,attribute,applyEdits,type XmlElement} from '../xml/index.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main';
const owners:Record<string,string>={sp:'nvSpPr',pic:'nvPicPr',grpSp:'nvGrpSpPr',graphicFrame:'nvGraphicFramePr',cxnSp:'nvCxnSpPr'};
function fail(message:string):never{throw new OoxmlError('PPTX_Z_ORDER_UNSUPPORTED',message);}
function id(n:XmlElement):number{const raw=attribute(n,'id'),v=raw&&/^\d+$/.test(raw)?Number(raw):NaN;if(!Number.isInteger(v)||v<1||v>2147483647)fail('Invalid graphical identity');return v;}
function one(parent:XmlElement,name:string):XmlElement{const rows=parent.children.filter(n=>n.namespaceURI===P&&n.localName===name);if(rows.length!==1)fail('Expected unique '+name);return rows[0]!;}
function selected(source:string,groupId?:number){
 const doc=parseXml(source);if(doc.root.namespaceURI!==P||doc.root.localName!=='sld')fail('Expected slide root');const tree=one(one(doc.root,'cSld'),'spTree'),seen=new Set<number>();for(const n of elements(doc,'cNvPr',P)){const value=id(n);if(seen.has(value))fail('Duplicate graphical identity');seen.add(value);}
 let parent=tree;
 if(groupId!==undefined){if(!Number.isInteger(groupId)||groupId<1||groupId>2147483647)fail('Invalid group ID');const groups=elements(doc,'grpSp',P).filter(n=>id(one(one(n,'nvGrpSpPr'),'cNvPr'))===groupId);if(groups.length!==1)fail('Missing unique group parent');parent=groups[0]!;let ancestor=parent.parent;while(ancestor!==tree){if(!ancestor||ancestor.namespaceURI!==P||ancestor.localName!=='grpSp')fail('Group outside shape-tree ancestry');ancestor=ancestor.parent;}}
 if(parent.children.length<2||parent.children[0]!.namespaceURI!==P||parent.children[0]!.localName!=='nvGrpSpPr'||parent.children[1]!.namespaceURI!==P||parent.children[1]!.localName!=='grpSpPr')fail('Invalid graphical parent metadata');one(one(parent,'nvGrpSpPr'),'cNvPr');one(parent,'grpSpPr');
 const nodes:XmlElement[]=[],ids:number[]=[];let cursor=parent.openEnd;
 for(let index=0;index<parent.children.length;index++){const n=parent.children[index]!;if(source.slice(cursor,n.start).trim())fail('Graphical parent lexical barrier');cursor=n.end;if(n.namespaceURI!==P)fail('Foreign graphical sibling');if(index<2)continue;if(n.localName==='extLst'){if(index!==parent.children.length-1)fail('Nonterminal extension metadata');continue;}const owner=owners[n.localName];if(!owner)fail('Unsupported graphical sibling');nodes.push(n);ids.push(id(one(one(n,owner),'cNvPr')));}
 if(source.slice(cursor,parent.closeStart).trim())fail('Graphical parent lexical barrier');return {nodes,ids};
}
export function getShapeOrderXml(source:string,groupId?:number):number[]{return [...selected(source,groupId).ids];}
export function reorderShapesXml(source:string,order:number[],groupId?:number):string{
 if(!Array.isArray(order)||Object.getPrototypeOf(order)!==Array.prototype||order.length>1000)fail('Order requires a bounded plain array');
 for(const k of Reflect.ownKeys(order))if(k!=='length'&&(typeof k!=='string'||!/^\d+$/.test(k)||String(Number(k))!==k||Number(k)>=order.length))fail('Unknown order property');
 const copy:number[]=[],seen=new Set<number>();for(let i=0;i<order.length;i++){const d=Object.getOwnPropertyDescriptor(order,String(i));if(!d||!('value'in d)||!Number.isInteger(d.value)||d.value<1||d.value>2147483647||seen.has(d.value))fail('Order requires unique plain identity values');seen.add(d.value);copy.push(d.value);}
 const {nodes,ids}=selected(source,groupId);if(copy.some(id=>!ids.includes(id)))fail('Selected identity is not a direct graphical sibling');const slots=ids.map((id,index)=>seen.has(id)?index:-1).filter(i=>i>=0);if(slots.every((slot,i)=>ids[slot]===copy[i]))return source;
 return applyEdits(source,slots.map((slot,i)=>{const target=nodes[slot]!,original=nodes[ids.indexOf(copy[i]!)]!;return {start:target.start,end:target.end,value:source.slice(original.start,original.end)};}));
}
