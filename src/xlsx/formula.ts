import {OoxmlError} from '../errors.ts';
import {parseA1Range,type A1Range} from './range.ts';
export interface FormulaReference extends A1Range {start:number;end:number;}
type Token={kind:'value'|'reference'|'name'|'operator'|'end';text:string;start:number;end:number;range?:A1Range};
const FUNCTIONS=new Set(['IF','SUM','LOG10','ABS','MIN','MAX','AVERAGE','COUNT','COUNTA','ROUND','AND','OR','NOT']);
const PRECEDENCE:Record<string,number>={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5};
function fail(message:string):never{throw new OoxmlError('xlsx-formula-unsupported',message);}
function limit(message:string):never{throw new OoxmlError('xlsx-formula-limit',message);}
/** Conservative static A1 grammar. No evaluation, names, dynamic references or workbook access. */
export function analyzeFormulaReferences(source:string):FormulaReference[]{
 if(typeof source!=='string'||!source.length||!source.isWellFormed())fail('Expected a well-formed nonempty formula string');
 if(source.length>1024*1024)limit('Formula exceeds 1 MiB input limit');
 if(/[\u0000-\u001f\u007f-\u009f]/.test(source))fail('Control characters are unsupported');
 const offsets=new Uint32Array(source.length+1);let byte=0;
 for(let i=0;i<source.length;){const code=source.codePointAt(i)!;offsets[i]=byte;const width=code>0xffff?2:1;byte+=code<0x80?1:code<0x800?2:code<0x10000?3:4;i+=width;offsets[i]=byte;}
 if(byte>1024*1024)limit('Formula exceeds 1 MiB UTF-8 input limit');
 const reference=/(?:(?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!)?\$?[A-Za-z]+\$?[1-9][0-9]*(?::\$?[A-Za-z]+\$?[1-9][0-9]*)?/uy;
 const name=/[\p{L}_][\p{L}\p{N}_.]*/uy,number=/(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?/y;
 const tokens:Token[]=[];let pos=0;
 const push=(token:Token)=>{if(tokens.length>=100000)limit('Formula exceeds 100000 tokens');tokens.push(token);};
 const match=(pattern:RegExp)=>{pattern.lastIndex=pos;return pattern.exec(source)?.[0];};
 const callFollows=(end:number)=>{while(source[end]===' ')end++;return source[end]==='(';};
 while(pos<source.length){
  if(source[pos]===' '){pos++;continue;}
  const start=pos;
  if(source[pos]==='"'){
   pos++;let closed=false;while(pos<source.length){if(source[pos++]!=='"')continue;if(source[pos]==='"'){pos++;continue;}closed=true;break;}if(!closed)fail('Unclosed string literal');push({kind:'value',text:source.slice(start,pos),start,end:pos});continue;
  }
  // A token such as LOG10 is a supported call name only when followed by '(';
  // it is otherwise treated as an A1 token and bounded by the direct-range parser.
  const identifier=match(name);
  if(identifier&&callFollows(pos+identifier.length)){pos+=identifier.length;push({kind:'name',text:identifier.toUpperCase(),start,end:pos});continue;}
  const ref=match(reference);
  if(ref){let range:A1Range;try{range=parseA1Range(ref);}catch{fail('Invalid or out-of-grid cell reference');}pos+=ref.length;push({kind:'reference',text:ref,start,end:pos,range});continue;}
  const numeric=match(number);if(numeric){pos+=numeric.length;push({kind:'value',text:numeric,start,end:pos});continue;}
  if(identifier){pos+=identifier.length;const text=identifier.toUpperCase();if(text!=='TRUE'&&text!=='FALSE')fail('Unresolved names are unsupported');push({kind:'value',text,start,end:pos});continue;}
  const pair=source.slice(pos,pos+2),operator=['<=','>=','<>'].includes(pair)?pair:source[pos]!;
  if(!['+','-','*','/','^','&','=','<','>','<=','>=','<>','%','(',')',','].includes(operator))fail('Unsupported formula token');
  pos+=operator.length;push({kind:'operator',text:operator,start,end:pos});
 }
 tokens.push({kind:'end',text:'',start:pos,end:pos});let cursor=0;const refs:FormulaReference[]=[];
 const current=()=>tokens[cursor]!;
 function expression(min:number,depth:number):void{
  if(depth>128)limit('Formula nesting exceeds 128');
  const token=current();
  if(token.text==='+'||token.text==='-'){cursor++;expression(6,depth+1);}
  else if(token.text==='('){cursor++;expression(1,depth+1);if(current().text!==')')fail('Unclosed expression');cursor++;}
  else if(token.kind==='name'){
   if(!FUNCTIONS.has(token.text))fail('Unsupported function');cursor++;if(current().text!=='(')fail('Expected function call');cursor++;
   expression(1,depth+1);while(current().text===','){cursor++;expression(1,depth+1);}if(current().text!==')')fail('Unclosed function');cursor++;
  }else if(token.kind==='value'||token.kind==='reference'){
   cursor++;if(token.range){if(refs.length>=10000)limit('Formula exceeds 10000 references');refs.push({...token.range,start:offsets[token.start]!,end:offsets[token.end]!});}
  }else fail('Expected expression operand');
  while(current().text==='%')cursor++;
  while(Object.hasOwn(PRECEDENCE,current().text)&&PRECEDENCE[current().text]!>=min){const precedence=PRECEDENCE[current().text]!;cursor++;expression(precedence+1,depth+1);}
 }
 if(current().text==='=')cursor++;
 expression(1,0);if(current().kind!=='end')fail('Unexpected trailing or adjacent expression');return refs;
}
