import {OoxmlError} from '../errors.ts';
import {escapeText} from '../xml/index.ts';
import {appendTextBox,textBoxRequest} from './text-box.ts';
import {patchShapeStyle} from './retained-style.ts';
import {addConnectorXml,type ConnectorEndpoint} from './connectors.ts';
import type {PictureGeometry} from './picture-add.ts';
export type DiagramNode={key:string;text:string};
export type DiagramEdge={from:string;to:string};
export type DiagramOptions={x:number;y:number;nodeWidth:number;nodeHeight:number;gap:number;direction:'row'|'column'};
export type DiagramNodeReceipt={key:string;shapeId:number;geometry:PictureGeometry};
export type DiagramEdgeReceipt=DiagramEdge&{shapeId:number;start:ConnectorEndpoint;end:ConnectorEndpoint};
export type DiagramReceipt={partName:string;nodes:DiagramNodeReceipt[];edges:DiagramEdgeReceipt[]};
function fail(message:string):never{throw new OoxmlError('PPTX_DIAGRAM_UNSUPPORTED',message);}
function plain(value:unknown,allowed:string[]):Record<string,unknown>{if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Diagram inputs require plain data');const copy:Record<string,unknown>={};for(const key of Reflect.ownKeys(value)){if(typeof key!=='string'||!allowed.includes(key))fail('Unknown diagram field');const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Diagram accessors are unsupported');copy[key]=d.value;}return copy;}
function list(value:unknown,min:number,max:number):unknown[]{if(!Array.isArray(value)||value.length<min||value.length>max)fail('Diagram list exceeds bounded size');const result:unknown[]=[];for(let i=0;i<value.length;i++){const d=Object.getOwnPropertyDescriptor(value,String(i));if(!d||!('value'in d))fail('Diagram list accessors are unsupported');result.push(d.value);}return result;}
function key(value:unknown):string{if(typeof value!=='string'||!value.trim()||value.length>64)fail('Diagram keys require nonempty bounded text');try{escapeText(value);}catch{fail('Invalid diagram key');}return value;}
function integer(value:unknown,min:number):number{if(typeof value!=='number'||!Number.isInteger(value)||value<min||value>2147483647)fail('Diagram layout requires bounded integers');return value;}
/** Plan/author a complete deterministic graph in XML before the package transaction. */
export function addDiagramXml(source:string,nodes:DiagramNode[],edges:DiagramEdge[],options:DiagramOptions){
 const inputs=list(nodes,1,32).map(n=>{const d=plain(n,['key','text']),k=key(d.key);if(typeof d.text!=='string'||d.text.length>4096)fail('Diagram labels require bounded text');try{escapeText(d.text);}catch{fail('Invalid diagram label');}return {key:k,text:d.text};}),seen=new Set<string>();for(const n of inputs){if(seen.has(n.key))fail('Duplicate diagram key');seen.add(n.key);}
 const pairs=new Set<string>(),links=list(edges,0,64).map(e=>{const d=plain(e,['from','to']),from=key(d.from),to=key(d.to),pair=JSON.stringify([from,to]);if(!seen.has(from)||!seen.has(to)||from===to||pairs.has(pair))fail('Invalid or duplicate diagram edge');pairs.add(pair);return {from,to};});
 const o=plain(options,['x','y','nodeWidth','nodeHeight','gap','direction']),x=integer(o.x,0),y=integer(o.y,0),width=integer(o.nodeWidth,1),height=integer(o.nodeHeight,1),gap=integer(o.gap,0);if(!['row','column'].includes(o.direction as string))fail('Unknown diagram direction');
 const geometries=inputs.map((_,i)=>{const g={x:x+(o.direction==='row'?i*(width+gap):0),y:y+(o.direction==='column'?i*(height+gap):0),width,height};integer(g.x+width,0);integer(g.y+height,0);return g;});
 let next=source;const nodeRows:DiagramNodeReceipt[]=[],edgeRows:DiagramEdgeReceipt[]=[];
 try{
  for(let i=0;i<inputs.length;i++){const n=inputs[i]!,g=geometries[i]!,result=appendTextBox(next,textBoxRequest(n.text,g,{name:'Diagram node '+n.key}));next=patchShapeStyle(result.xml,result.shapeId,{fill:'F2F2F2',lineColor:'336699',lineWidth:12700});nodeRows.push({key:n.key,shapeId:result.shapeId,geometry:{...g}});}
  for(const edge of links){const start={shapeId:nodeRows.find(n=>n.key===edge.from)!.shapeId,site:o.direction==='row'?3:2},end={shapeId:nodeRows.find(n=>n.key===edge.to)!.shapeId,site:o.direction==='row'?1:0},result=addConnectorXml(next,start,end,{name:`Diagram edge ${edge.from} -> ${edge.to}`,color:'336699',width:12700});next=result.xml;edgeRows.push({...edge,shapeId:result.shapeId,start,end});}
 }catch(error){if(error instanceof OoxmlError&&error.code!=='PPTX_ID_EXHAUSTED')fail(error.message);throw error;}
 return {xml:next,nodes:nodeRows,edges:edgeRows};
}
