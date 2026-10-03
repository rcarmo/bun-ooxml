import {OoxmlError} from '../errors.ts';
import {parseXml,elements,applyEdits} from '../xml/index.ts';
import {addAutoShapeXml} from './autoshapes.ts';
import type {TextBoxGeometry} from './text-box.ts';
export type FreeformCommand={op:'move'|'line';x:number;y:number}|{op:'close'};
export type FreeformPath={width:number;height:number;commands:FreeformCommand[]};
export type FreeformOptions={name?:string;fill?:string;lineColor?:string;lineWidth?:number};
export type FreeformReceipt={shapeId:number;partName:string;geometry:TextBoxGeometry;path:FreeformPath};
const A='http://schemas.openxmlformats.org/drawingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main';
function fail(message:string):never{throw new OoxmlError('PPTX_FREEFORM_UNSUPPORTED',message);}
function plain(value:unknown,allowed:string[]):Record<string,unknown>{if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Freeform requires plain data');const result:Record<string,unknown>={};for(const key of Reflect.ownKeys(value)){if(typeof key!=='string'||!allowed.includes(key))fail('Unknown freeform field');const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Freeform accessors are unsupported');result[key]=d.value;}return result;}
function integer(value:unknown,min:number,max=2147483647):number{if(typeof value!=='number'||!Number.isInteger(value)||value<min||value>max)fail('Freeform requires bounded integer dimensions/points');return value;}
export function addFreeformXml(source:string,geometry:TextBoxGeometry,path:FreeformPath,options:FreeformOptions={}){
 const p=plain(path,['width','height','commands']),width=integer(p.width,1),height=integer(p.height,1),o=plain(options,['name','fill','lineColor','lineWidth']),fill=Object.hasOwn(o,'fill')?o.fill:'none';
 if(!Array.isArray(p.commands)||p.commands.length<2||p.commands.length>256)fail('Freeform requires 2–256 commands');
 const commands:FreeformCommand[]=[],vertices=new Set<string>();let active=false,lines=0,closed=false,last='';
 function finish(){if(active&&(lines<1||fill!=='none'&&!closed))fail('Incomplete/degenerate freeform subpath');}
 for(let i=0;i<p.commands.length;i++){
  const d=Object.getOwnPropertyDescriptor(p.commands,String(i));if(!d||!('value'in d))fail('Freeform command accessors are unsupported');const c=plain(d.value,['op','x','y']);
  if(c.op==='close'){if(Object.keys(c).length!==1||!active||closed||lines<2||vertices.size<3)fail('Close requires a polygonal subpath');commands.push({op:'close'});closed=true;continue;}
  if(c.op!=='move'&&c.op!=='line')fail('Unsupported freeform command');const x=integer(c.x,0,width),y=integer(c.y,0,height),point=x+','+y;
  if(c.op==='move'){finish();active=true;closed=false;lines=0;vertices.clear();last=point;vertices.add(point);}
  else{if(!active||closed||point===last)fail('Line requires active nonduplicate subpath');lines++;last=point;vertices.add(point);}
  commands.push({op:c.op,x,y});
 }
 finish();
 try{
  const added=addAutoShapeXml(source,'rect',geometry,{...o,fill:fill as string} as FreeformOptions);const doc=parseXml(added.xml),shape=elements(doc,'sp',P).find(n=>n.children.find(c=>c.localName==='nvSpPr'&&c.namespaceURI===P)?.children.some(c=>c.localName==='cNvPr'&&c.namespaceURI===P&&c.attributes.id===String(added.shapeId)))!,preset=elements(shape,'prstGeom',A)[0]!,textBody=shape.children.find(n=>n.localName==='txBody'&&n.namespaceURI===P)!;
  const body=commands.map(c=>c.op==='close'?'<a:close/>':`<a:${c.op==='move'?'moveTo':'lnTo'}><a:pt x="${c.x}" y="${c.y}"/></a:${c.op==='move'?'moveTo':'lnTo'}>`).join(''),custom=`<a:custGeom xmlns:a="${A}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${width}" h="${height}" fill="${fill==='none'?'none':'norm'}" stroke="1" extrusionOk="0">${body}</a:path></a:pathLst></a:custGeom>`;
  let next=applyEdits(added.xml,[{start:preset.start,end:preset.end,value:custom},{start:textBody.start,end:textBody.end,value:''}]);
  if(!Object.hasOwn(o,'name')){const id=elements(shape,'cNvPr',P)[0]!;next=applyEdits(next,[{start:id.start,end:id.openEnd,value:next.slice(id.start,id.openEnd).replace(`name="AutoShape ${added.shapeId}"`,`name="Freeform ${added.shapeId}"`)}]);}
  return {xml:next,shapeId:added.shapeId,geometry:{...added.geometry},path:{width,height,commands:commands.map(c=>({...c}))}};
 }catch(error){if(error instanceof OoxmlError&&error.code!=='PPTX_ID_EXHAUSTED')fail(error.message);throw error;}
}
