import {OoxmlError} from '../errors.ts';
export type A1RangeAxis='cell'|'row'|'column';
export interface A1Coordinate {row:number;column:number;rowAbsolute:boolean;columnAbsolute:boolean;}
export interface A1Range {sheet:string;axis:A1RangeAxis;first:A1Coordinate;last:A1Coordinate;}
const fail=(message:string):never=>{throw new OoxmlError('xlsx-range-unsupported',message);};
const MAX_ROW=1_048_576,MAX_COL=16_384,MAX_INPUT=1_048_576;
/** Parse only a direct A1 reference, never an expression or a workbook lookup.
 * Reversed endpoints retain their order. Whole axes have one zero coordinate.
 */
export function parseA1Range(source:string):A1Range {
 if(typeof source!=='string'||!source.length||source.length>MAX_INPUT)fail('Expected a bounded nonempty direct A1 range');
 if(!source.isWellFormed()||/[\u0000-\u001f\u007f]/.test(source))fail('Invalid character in direct range');
 if(new TextEncoder().encode(source).length>MAX_INPUT)fail('Direct range exceeds UTF-8 byte limit');
 let sheet='',range=source;
 if(source.startsWith("'")){
  let i=1,name='',closed=false;
  while(i<source.length){const c=source[i++]!;if(c!=="'"){name+=c;continue;}if(source[i]==="'"){name+="'";i++;continue;}closed=true;break;}
  if(!closed||source[i]!=='!')fail('Unclosed or unqualified quoted sheet');sheet=name;range=source.slice(i+1);
 }else if(source.includes('!')){
  const at=source.indexOf('!');sheet=source.slice(0,at);range=source.slice(at+1);
  if(!/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(sheet))fail('Unsupported unquoted sheet qualifier');
 }
 if(sheet){if(/[\[\]:*?/\\]/.test(sheet))fail('Unsupported external or three-dimensional sheet qualifier');}
 else if(range!==source)fail('Empty sheet qualifier');
 const pieces=range.split(':');if(pieces.length>2||pieces.some(p=>!p))fail('Expected one or two direct endpoints');
 const first=endpoint(pieces[0]!);if(pieces.length===1&&first.axis!=='cell')fail('Whole axes require two endpoints');
 const last=pieces.length===2?endpoint(pieces[1]!):{axis:first.axis,value:{...first.value}};
 if(first.axis!==last.axis)fail('Mixed endpoint axes');
 return {sheet,axis:first.axis,first:first.value,last:last.value};
}
function endpoint(text:string):{axis:A1RangeAxis;value:A1Coordinate}{
 let m=/^(\$?)([A-Za-z]+)(\$?)([1-9][0-9]*)$/.exec(text);
 if(m)return {axis:'cell',value:{row:row(m[4]!),column:column(m[2]!),rowAbsolute:m[3]==='$',columnAbsolute:m[1]==='$'}};
 m=/^(\$?)([A-Za-z]+)$/.exec(text);if(m)return {axis:'column',value:{row:0,column:column(m[2]!),rowAbsolute:false,columnAbsolute:m[1]==='$'}};
 m=/^(\$?)([1-9][0-9]*)$/.exec(text);if(m)return {axis:'row',value:{row:row(m[2]!),column:0,rowAbsolute:m[1]==='$',columnAbsolute:false}};
 return fail('Unsupported direct range endpoint');
}
function row(text:string){const n=Number(text);if(!Number.isSafeInteger(n)||n<1||n>MAX_ROW)fail('Row outside supported worksheet grid');return n;}
function column(text:string){let n=0;for(const c of text.toUpperCase()){n=n*26+c.charCodeAt(0)-64;if(n>MAX_COL)fail('Column outside supported worksheet grid');}return n;}
