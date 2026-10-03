import {OoxmlError} from '../errors.ts';
import {parseXml,elements,attribute,applyEdits,type XmlElement} from '../xml/index.ts';
import {retainedFormattingXml} from './formatting.ts';
export type GroupTransform={x:number;y:number;width:number;height:number;childX:number;childY:number;childWidth:number;childHeight:number;rotation:number;flipH:boolean;flipV:boolean};
export type GroupTransformPatch=Partial<GroupTransform>;
export type GroupPoint={x:number;y:number};
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',N='http://www.w3.org/2000/xmlns/',MAX=2147483647;
const keys=['x','y','width','height','childX','childY','childWidth','childHeight','rotation','flipH','flipV'] as const;
function fail(message:string):never{throw new OoxmlError('PPTX_GROUP_UNSUPPORTED',message);}
function plain(value:unknown,allowed:readonly string[]):Record<string,unknown>{if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Group frame requires plain data');const copy:Record<string,unknown>={};for(const k of Reflect.ownKeys(value)){if(typeof k!=='string'||!allowed.includes(k))fail('Unknown group frame field');const d=Object.getOwnPropertyDescriptor(value,k)!;if(!('value'in d))fail('Group frame accessors are unsupported');copy[k]=d.value;}return copy;}
function scalar(key:string,value:unknown):number|boolean{if(key==='flipH'||key==='flipV'){if(typeof value!=='boolean')fail('Group flips must be Boolean');return value;}const min=['x','y','childX','childY'].includes(key)?-2147483648:key==='rotation'?0:1,max=key==='rotation'?21599999:MAX;if(typeof value!=='number'||!Number.isInteger(value)||value<min||value>max)fail('Group frame requires bounded integer values');return value;}
function frame(value:unknown):GroupTransform{const copy=plain(value,keys);for(const k of keys)copy[k]=scalar(k,copy[k]);return copy as GroupTransform;}
function one(parent:XmlElement,name:string,ns:string):XmlElement{const rows=parent.children.filter(n=>n.localName===name&&n.namespaceURI===ns);if(rows.length!==1)fail('Expected unique '+name);return rows[0]!;}
function attrs(n:XmlElement,allowed:string[],required:string[]=[]){const k=Object.keys(n.attributes).filter(k=>n.attributeNamespaces[k]!==N);if(k.some(k=>n.attributeNamespaces[k]!==''||!allowed.includes(k))||required.some(name=>!k.includes(name)))fail('Unsupported group transform attributes');}
function integer(n:XmlElement,k:string):number{const v=attribute(n,k);if(v===undefined||!/^[-+]?\d+$/.test(v))fail('Missing integer group attribute');return Number(v);}
function boolean(n:XmlElement,k:string):boolean{const v=attribute(n,k);if(v===undefined||v==='0'||v==='false')return false;if(v==='1'||v==='true')return true;return fail('Invalid group flip');}
function selected(source:string,id:number){
 if(!Number.isInteger(id)||id<1||id>MAX)fail('Invalid group ID');const doc=parseXml(source);if(doc.root.namespaceURI!==P||doc.root.localName!=='sld')fail('Expected slide root');const tree=one(one(doc.root,'cSld',P),'spTree',P),seen=new Set<number>();
 for(const n of elements(doc,'cNvPr',P)){const v=integer(n,'id');if(v<1||v>MAX||seen.has(v))fail('Duplicate or invalid slide identity');seen.add(v);}
 const groups=elements(doc,'grpSp',P).filter(n=>Number(attribute(one(one(n,'nvGrpSpPr',P),'cNvPr',P),'id'))===id);if(groups.length!==1)fail('Missing unique group');const group=groups[0]!;
 let ancestor=group.parent;while(ancestor!==tree){if(!ancestor||ancestor.namespaceURI!==P||ancestor.localName!=='grpSp')fail('Group outside direct shape-tree ancestry');ancestor=ancestor.parent;}
 const transform=one(one(group,'grpSpPr',P),'xfrm',A);attrs(transform,['rot','flipH','flipV']);const names=['off','ext','chOff','chExt'];if(transform.children.length!==4||transform.children.some((n,i)=>n.namespaceURI!==A||n.localName!==names[i]))fail('Unsupported group transform child order');
 let cursor=transform.openEnd;for(const n of transform.children){if(source.slice(cursor,n.start).trim()||n.children.length||(!n.selfClosing&&source.slice(n.openEnd,n.closeStart).trim()))fail('Group transform lexical barrier');cursor=n.end;const allowed=n.localName==='off'||n.localName==='chOff'?['x','y']:['cx','cy'];attrs(n,allowed,allowed);}if(source.slice(cursor,transform.closeStart).trim())fail('Group transform lexical barrier');
 const [off,size,childOff,childSize]=transform.children;
 const value=frame({x:integer(off!,'x'),y:integer(off!,'y'),width:integer(size!,'cx'),height:integer(size!,'cy'),childX:integer(childOff!,'x'),childY:integer(childOff!,'y'),childWidth:integer(childSize!,'cx'),childHeight:integer(childSize!,'cy'),rotation:attribute(transform,'rot')===undefined?0:integer(transform,'rot'),flipH:boolean(transform,'flipH'),flipV:boolean(transform,'flipV')});return {transform,value};
}
export function getGroupTransformXml(source:string,id:number):GroupTransform{return {...selected(source,id).value};}
export function patchGroupTransformXml(source:string,id:number,patch:GroupTransformPatch):string{
 const change=plain(patch,keys);for(const k of Object.keys(change))change[k]=scalar(k,change[k]);const {transform,value}=selected(source,id),next=frame({...value,...change});
 const specs:[XmlElement,Record<string,string>][]=[[transform,{}],...transform.children.map(n=>[n,{}] as [XmlElement,Record<string,string>])];
 const targets:Record<string,[number,string]>={x:[1,'x'],y:[1,'y'],width:[2,'cx'],height:[2,'cy'],childX:[3,'x'],childY:[3,'y'],childWidth:[4,'cx'],childHeight:[4,'cy'],rotation:[0,'rot'],flipH:[0,'flipH'],flipV:[0,'flipV']};
 for(const k of keys)if(next[k]!==value[k]){const [index,attr]=targets[k]!;specs[index]![1][attr]=typeof next[k]==='boolean'?next[k]?'1':'0':String(next[k]);}
 const edits=specs.filter(([,p])=>Object.keys(p).length).map(([n,p])=>({start:n.start,end:n.openEnd,value:retainedFormattingXml.openTag(source,n,p)}));return applyEdits(source,edits);
}
function point(value:unknown):GroupPoint{const p=plain(value,['x','y']);for(const k of ['x','y'])if(typeof p[k]!=='number'||!Number.isFinite(p[k])||Math.abs(p[k] as number)>MAX)fail('Point exceeds finite mapping bounds');return p as GroupPoint;}
function trig(rotation:number):[number,number]{const quarter=rotation/5400000;if(Number.isInteger(quarter))return ([[1,0],[0,1],[-1,0],[0,-1]] as [number,number][])[quarter]!;const angle=rotation/60000*Math.PI/180;return [Math.cos(angle),Math.sin(angle)];}
export function mapGroupPoint(transform:GroupTransform,input:GroupPoint):GroupPoint{
 const t=frame(transform),p=point(input),[c,s]=trig(t.rotation),centre=point({x:t.x+t.width/2,y:t.y+t.height/2}),cx=centre.x,cy=centre.y,scaled=point({x:t.x+(p.x-t.childX)*t.width/t.childWidth,y:t.y+(p.y-t.childY)*t.height/t.childHeight});let dx=scaled.x-cx,dy=scaled.y-cy;if(t.flipH)dx=-dx;if(t.flipV)dy=-dy;const reflected=point({x:cx+dx,y:cy+dy});dx=reflected.x-cx;dy=reflected.y-cy;return point({x:cx+c*dx-s*dy,y:cy+s*dx+c*dy});
}
export function unmapGroupPoint(transform:GroupTransform,input:GroupPoint):GroupPoint{
 const t=frame(transform),p=point(input),[c,s]=trig(t.rotation),centre=point({x:t.x+t.width/2,y:t.y+t.height/2}),cx=centre.x,cy=centre.y,dx=p.x-cx,dy=p.y-cy,unrotated=point({x:cx+c*dx+s*dy,y:cy-s*dx+c*dy});let x=unrotated.x-cx,y=unrotated.y-cy;if(t.flipH)x=-x;if(t.flipV)y=-y;const reflected=point({x:cx+x,y:cy+y});return point({x:t.childX+(reflected.x-t.x)*t.childWidth/t.width,y:t.childY+(reflected.y-t.y)*t.childHeight/t.height});
}
function chain(value:GroupTransform[]):GroupTransform[]{if(!Array.isArray(value)||value.length<1||value.length>16)fail('Mapping chain requires 1–16 frames');const result:GroupTransform[]=[];for(let i=0;i<value.length;i++){const d=Object.getOwnPropertyDescriptor(value,String(i));if(!d||!('value'in d))fail('Mapping chain accessors are unsupported');result.push(frame(d.value));}return result;}
export function mapGroupPointChain(transforms:GroupTransform[],input:GroupPoint):GroupPoint{return chain(transforms).reverse().reduce((p,t)=>mapGroupPoint(t,p),point(input));}
export function unmapGroupPointChain(transforms:GroupTransform[],input:GroupPoint):GroupPoint{return chain(transforms).reduce((p,t)=>unmapGroupPoint(t,p),point(input));}
