import {OoxmlError} from '../errors.ts';
import {analyzeFormulaReferences,analyzeProfileFormulaReferences} from './formula.ts';
import type {A1Coordinate} from './range.ts';
export interface FormulaInsertion {axis:'row'|'column';at:number;count:number;sheet:string;}
function fail(message:string):never{throw new OoxmlError('xlsx-formula-remap-unsupported',message);}
function sheetName(value:unknown):string{
 if(typeof value!=='string'||!value.trim()||value.length>255||!value.isWellFormed()||/[\u0000-\u001f\u007f-\u009f\[\]:*?/\\]/.test(value))fail('Expected a nonblank bounded sheet name without external or 3D syntax');return value;
}
function insertion(value:FormulaInsertion):FormulaInsertion{
 if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Insertion must be plain data');
 const copy:Record<string,unknown>={};for(const key of Reflect.ownKeys(value)){
  if(typeof key!=='string'||!['axis','at','count','sheet'].includes(key))fail('Unknown insertion field');const descriptor=Object.getOwnPropertyDescriptor(value,key)!;if(!('value' in descriptor))fail('Insertion fields cannot be accessors');copy[key]=descriptor.value;
 }
 if(copy.axis!=='row'&&copy.axis!=='column')fail('Expected row or column insertion');const max=copy.axis==='row'?1048576:16384;
 for(const key of ['at','count'])if(typeof copy[key]!=='number'||!Number.isSafeInteger(copy[key])||copy[key]<1||copy[key]>max)fail('Insertion coordinate and count must be positive integers within the axis limit');
 return {axis:copy.axis,at:copy.at as number,count:copy.count as number,sheet:sheetName(copy.sheet)};
}
// Unicode simple lowercase is codepoint-wise and never full case folding.
function simpleLower(value:string):string{return [...value].map(c=>c==='İ'?'i':c.toLowerCase()).join('');}
function cell(c:A1Coordinate):string{let column=c.column,name='';while(column){column--;name=String.fromCharCode(65+column%26)+name;column=Math.floor(column/26);}return (c.columnAbsolute?'$':'')+name+(c.rowAbsolute?'$':'')+c.row;}
/** Rewrite supported static references after one insertion; never mutates a workbook. */
export function insertFormulaReferences(source:string,contextSheet:string,change:FormulaInsertion):string{return remap(source,contextSheet,change,false);}
export function insertProfileFormulaReferences(source:string,contextSheet:string,change:FormulaInsertion):string{return remap(source,contextSheet,change,true);}
function remap(source:string,contextSheet:string,change:FormulaInsertion,profile:boolean):string{
 const context=sheetName(contextSheet),edit=insertion(change),references=(profile?analyzeProfileFormulaReferences:analyzeFormulaReferences)(source),encoder=new TextEncoder(),decoder=new TextDecoder(),bytes=encoder.encode(source);
 const patches:Array<{start:number;end:number;value:string}>=[];let size=bytes.length;
 for(const reference of references){
  if(profile?simpleLower(reference.sheet||context)!==simpleLower(edit.sheet):(reference.sheet||context).toLowerCase()!==edit.sheet.toLowerCase())continue;
  const first={...reference.first},last={...reference.last},key=edit.axis==='row'?'row':'column',max=edit.axis==='row'?1048576:16384;let changed=false;
  for(const point of [first,last])if(point[key]>=edit.at){if(point[key]>max-edit.count)fail('Insertion moves a referenced coordinate outside the grid');point[key]+=edit.count;changed=true;}
  if(!changed)continue;
  const original=decoder.decode(bytes.subarray(reference.start,reference.end)),bang=original.lastIndexOf('!'),qualifier=original.slice(0,bang+1),range=original.slice(bang+1);
  const value=qualifier+cell(first)+(range.includes(':')?':'+cell(last):'');
  size+=encoder.encode(value).length-(reference.end-reference.start);if(size>1024*1024)throw new OoxmlError('xlsx-formula-limit','Remapped formula exceeds 1 MiB UTF-8 output limit');
  patches.push({start:reference.start,end:reference.end,value});
 }
 if(!patches.length)return source;
 const parts:string[]=[];let cursor=0;for(const patch of patches){parts.push(decoder.decode(bytes.subarray(cursor,patch.start)),patch.value);cursor=patch.end;}parts.push(decoder.decode(bytes.subarray(cursor)));return parts.join('');
}
