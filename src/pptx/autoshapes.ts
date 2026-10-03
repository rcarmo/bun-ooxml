import {OoxmlError} from '../errors.ts';
import {parseXml,elements,applyEdits} from '../xml/index.ts';
import {appendTextBox,textBoxRequest,type TextBoxGeometry} from './text-box.ts';
import {patchShapeStyle} from './retained-style.ts';
export type AutoShapePreset='rect'|'ellipse'|'triangle'|'diamond'|'roundRect';
export type AutoShapeOptions={name?:string;text?:string;adjustments?:{adj?:number};fill?:string;lineColor?:string;lineWidth?:number};
export type AutoShapeReceipt={shapeId:number;partName:string;preset:AutoShapePreset;geometry:TextBoxGeometry;adjustments:Record<string,number>};
const A='http://schemas.openxmlformats.org/drawingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main';
function fail(message:string):never{throw new OoxmlError('PPTX_AUTOSHAPE_UNSUPPORTED',message);}
function plain(value:unknown,allowed:string[]):Record<string,unknown>{if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('AutoShape options require plain data');const copy:Record<string,unknown>={};for(const k of Reflect.ownKeys(value)){if(typeof k!=='string'||!allowed.includes(k))fail('Unknown AutoShape field');const d=Object.getOwnPropertyDescriptor(value,k)!;if(!('value'in d))fail('AutoShape accessors are unsupported');copy[k]=d.value;}return copy;}
/** Reuse native editable text/style, replacing only freshly authored preset geometry. */
export function addAutoShapeXml(source:string,preset:AutoShapePreset,geometry:TextBoxGeometry,options:AutoShapeOptions={}){
 if(!['rect','ellipse','triangle','diamond','roundRect'].includes(preset))fail('Unsupported AutoShape preset');
 const o=plain(options,['name','text','adjustments','fill','lineColor','lineWidth']),adjustment=plain(Object.hasOwn(o,'adjustments')?o.adjustments:{},preset==='roundRect'?['adj']:[]),values:Record<string,number>={};
 for(const key of ['name','text','fill','lineColor','lineWidth'])if(Object.hasOwn(o,key)&&(o[key]===undefined||o[key]===null))fail('AutoShape option must be explicit data');
 if(preset==='roundRect'){const v=Object.hasOwn(adjustment,'adj')?adjustment.adj:16667;if(typeof v!=='number'||!Number.isInteger(v)||v<0||v>50000)fail('Round rectangle adjustment must be in [0,50000]');values.adj=v;}
 const text=Object.hasOwn(o,'text')?o.text:'';if(typeof text!=='string'||text.length>4096)fail('AutoShape text requires bounded XML text');
 try{
  const request=textBoxRequest(text,geometry,o.name!==undefined?{name:o.name as string}:{}),added=appendTextBox(source,request);let next=patchShapeStyle(added.xml,added.shapeId,{fill:(Object.hasOwn(o,'fill')?o.fill:'F2F2F2') as string,lineColor:(Object.hasOwn(o,'lineColor')?o.lineColor:'336699') as string,lineWidth:(Object.hasOwn(o,'lineWidth')?o.lineWidth:12700) as number});
  const doc=parseXml(next),shape=elements(doc,'sp',P).find(n=>n.children.find(c=>c.namespaceURI===P&&c.localName==='nvSpPr')?.children.some(c=>c.namespaceURI===P&&c.localName==='cNvPr'&&c.attributes.id===String(added.shapeId)))!,pr=elements(shape,'prstGeom',A)[0]!;
  const guides=Object.entries(values).map(([name,value])=>`<a:gd name="${name}" fmla="val ${value}"/>`).join(''),value=`<a:prstGeom xmlns:a="${A}" prst="${preset}"><a:avLst>${guides}</a:avLst></a:prstGeom>`;
  const nv=shape.children.find(n=>n.localName==='nvSpPr'&&n.namespaceURI===P)!,nonvisual=nv.children.find(n=>n.localName==='cNvSpPr'&&n.namespaceURI===P)!;
  next=applyEdits(next,[{start:pr.start,end:pr.end,value},{start:nonvisual.start,end:nonvisual.end,value:'<p:cNvSpPr/>'}]);
  if(o.name===undefined){const props=shape.children.find(n=>n.localName==='nvSpPr'&&n.namespaceURI===P)!.children.find(n=>n.localName==='cNvPr'&&n.namespaceURI===P)!;next=applyEdits(next,[{start:props.start,end:props.openEnd,value:next.slice(props.start,props.openEnd).replace(`name="TextBox ${added.shapeId}"`,`name="AutoShape ${added.shapeId}"`)}]);}
  return {xml:next,shapeId:added.shapeId,preset,geometry:{...request.geometry},adjustments:{...values}};
 }catch(error){if(error instanceof OoxmlError&&error.code!=='PPTX_ID_EXHAUSTED')fail(error.message);throw error;}
}
