import {beforeUniformApi18Feature} from '../helpers/uniform-api18-history.ts';
import {test,expect} from 'bun:test';
import * as api from '../../src/index.ts';
const {insertFormulaReferences,analyzeFormulaReferences}=api;

test('static insertions rewrite exact shared expressions without touching literals or unrelated sheets',()=>{
 for(const[axis,at,count,sheet,source,expected]of [
  ['row',2,1,'Main','IF(A1="A2",A2,Other!A2)','IF(A1="A2",A3,Other!A2)'],
  ['column',2,1,'Main',"'O''Brien'!$b$2 + Main!c1","'O''Brien'!$b$2 + Main!D1"],
  ['row',3,2,'Main','SUM(A5:A1)','SUM(A7:A1)'],
  ['row',3,2,'Main','a1 + Other!b2','a1 + Other!b2'],
  ['row',2,1,'main','Main!A1:A3','Main!A1:A4'],
 ] as const)expect(insertFormulaReferences(source,'Main',{axis,at,count,sheet})).toBe(expected);
});

test('absolute coordinates shift structurally while relative flags and reversed endpoints remain independent',()=>{
 expect(insertFormulaReferences('$b2:C$4','Main',{axis:'row',at:3,count:2,sheet:'Main'})).toBe('$B2:C$6');
 expect(insertFormulaReferences('$B$2:C4','Main',{axis:'column',at:2,count:2,sheet:'Main'})).toBe('$D$2:E4');
 expect(insertFormulaReferences('D9:B3','Main',{axis:'row',at:5,count:3,sheet:'Main'})).toBe('D12:B3');
 expect(insertFormulaReferences('A1:A1','Main',{axis:'row',at:1,count:1,sheet:'Main'})).toBe('A2:A2');
 const out=insertFormulaReferences("'O''Brien'!$B2:C$4",'Other',{axis:'row',at:1,count:1,sheet:"o'brien"}),ref=analyzeFormulaReferences(out)[0]!;
 expect(out).toBe("'O''Brien'!$B3:C$5");expect(ref.first).toEqual({row:3,column:2,rowAbsolute:false,columnAbsolute:true});expect(ref.last).toEqual({row:5,column:3,rowAbsolute:true,columnAbsolute:false});
});

test('UTF8 edits preserve quoted qualifier spelling and every non-reference byte across multiple changes',()=>{
 const source=' =IF("雪😀A2"="A2",\'α ! sheet\'!$b$2 + Other!c3,\'α ! sheet\'!D4) ';
 expect(insertFormulaReferences(source,'Elsewhere',{axis:'column',at:2,count:1,sheet:'Α ! SHEET'})).toBe(' =IF("雪😀A2"="A2",\'α ! sheet\'!$C$2 + Other!c3,\'α ! sheet\'!E4) ');
 expect(insertFormulaReferences("'3'!a1 + 'O''Brien'!a1",'Main',{axis:'row',at:1,count:9,sheet:'3'})).toBe("'3'!A10 + 'O''Brien'!a1");
});

test('unaffected references and no-reference expressions retain exact spelling but still undergo analysis',()=>{
 for(const source of [' a1 + Other!b2 ', 'SUM("A1",12.00,TRUE)', '"雪😀"&"A1"', 'Main!a1:Z9'])expect(insertFormulaReferences(source,'Main',{axis:'row',at:100,count:1,sheet:'main'})).toBe(source);
 expect(()=>insertFormulaReferences('A1+INDIRECT(B2)','Main',{axis:'row',at:100,count:1,sheet:'Other'})).toThrow();
 expect(()=>insertFormulaReferences('A1+','Main',{axis:'row',at:100,count:1,sheet:'Other'})).toThrow();
});

test('late row and column overflow throw with no partial replacement and unchanged options',()=>{
 for(const[source,axis]of [['A1+A1048576','row'],['A1+XFD1','column'],['A1:A1048576','row'],['XFD1:A1','column']] as const){const change={axis,at:1,count:1,sheet:'Main'},before=JSON.stringify(change);let result:string|undefined;expect(()=>{result=insertFormulaReferences(source,'Main',change);}).toThrow(expect.objectContaining({code:'xlsx-formula-remap-unsupported'}));expect(result).toBeUndefined();expect(JSON.stringify(change)).toBe(before);}
 expect(insertFormulaReferences('Other!XFD1048576+A1','Main',{axis:'row',at:1,count:1,sheet:'Main'})).toBe('Other!XFD1048576+A2');
});

test('invalid arguments, inherited fields and executable patches refuse without invoking accessors',()=>{
 let called=false;const getter=Object.defineProperty({axis:'row',at:1,count:1},'sheet',{enumerable:true,get(){called=true;return 'Main';}});
 const good={axis:'row',at:1,count:1,sheet:'Main'};
 for(const change of [null,[],{},Object.create(good),getter,{...good,extra:1},{...good,[Symbol('x')]:1},{...good,axis:'bad'},{...good,at:0},{...good,at:1.5},{...good,at:1048577},{...good,count:0},{...good,count:NaN},{...good,count:1.5},{...good,count:Number.MAX_SAFE_INTEGER},{...good,sheet:''},{...good,sheet:'  '},{...good,sheet:'x\n'},{...good,sheet:'[Book]Main'},{...good,sheet:'\ud800'}])expect(()=>insertFormulaReferences('A1','Main',change as api.FormulaInsertion)).toThrow();
 for(const context of ['', ' ', 'bad\u0000', 'A:B', 'x'.repeat(256),null])expect(()=>insertFormulaReferences('A1',context as string,good as api.FormulaInsertion)).toThrow();expect(called).toBe(false);
 const plain=Object.assign(Object.create(null),good);expect(insertFormulaReferences('A1','Main',plain)).toBe('A2');expect(plain).toEqual(good);
});

test('grid edge insertions obey coordinate thresholds and bounded counts',()=>{
 expect(insertFormulaReferences('A1048575','Main',{axis:'row',at:1048575,count:1,sheet:'Main'})).toBe('A1048576');
 expect(insertFormulaReferences('XFC1','Main',{axis:'column',at:16383,count:1,sheet:'Main'})).toBe('XFD1');
 expect(insertFormulaReferences('A1','Main',{axis:'row',at:1048576,count:1048576,sheet:'Main'})).toBe('A1');
 expect(()=>insertFormulaReferences('XFD1','Main',{axis:'column',at:16384,count:1,sheet:'Main'})).toThrow();
 expect(()=>insertFormulaReferences('A1','Main',{axis:'column',at:16385,count:1,sheet:'Main'})).toThrow();
 expect(()=>insertFormulaReferences('A1','Main',{axis:'column',at:1,count:16385,sheet:'Main'})).toThrow();
});

test('finite 288-expression matrix reparses slices, preserves unaffected expressions and shifts wrapped spans',()=>{
 let total=0;const encoder=new TextEncoder(),decoder=new TextDecoder();for(const prefix of ['', 'Main!',"'Input Data'!"])for(const a of ['A1','$B2','C$3','$XFD$9'])for(const b of ['A1','$B2','C$3','$XFD$9'])for(const op of ['+','-','*','/','&','>=']){const source=prefix+a+op+prefix+b,refs=analyzeFormulaReferences(source),wrapped=analyzeFormulaReferences('SUM('+source+')');expect(refs).toHaveLength(2);expect(insertFormulaReferences(source,'Main',{axis:'row',at:100,count:1,sheet:'Main'})).toBe(source);expect(wrapped).toHaveLength(2);for(let i=0;i<2;i++){const ref=refs[i]!,token=decoder.decode(encoder.encode(source).slice(ref.start,ref.end)),one=analyzeFormulaReferences(token);expect(one).toHaveLength(1);expect(one[0]).toEqual({...ref,start:0,end:encoder.encode(token).length});expect(wrapped[i]).toEqual({...ref,start:ref.start+4,end:ref.end+4});}total++;}expect(total).toBe(288);
});

test('output growth refuses at UTF8 byte bound without returning a partial formula',()=>{
 const source='"'+'x'.repeat(1024*1024-9)+'"+A1+A9';expect(new TextEncoder().encode(source)).toHaveLength(1024*1024-1);let output:string|undefined;expect(()=>{output=insertFormulaReferences(source,'Main',{axis:'row',at:1,count:999,sheet:'Main'});}).toThrow(expect.objectContaining({code:'xlsx-formula-limit'}));expect(output).toBeUndefined();expect(source.endsWith('+A1+A9')).toBe(true);
});

test('qualified single cells remain single and colon/bang characters in quoted sheet names cannot alter range shape',()=>{
 expect(insertFormulaReferences("'Wow!'!a1",'Other',{axis:'row',at:1,count:1,sheet:'wow!'})).toBe("'Wow!'!A2");
 expect(insertFormulaReferences("'Main'!a1:a1",'Other',{axis:'column',at:1,count:1,sheet:'MAIN'})).toBe("'Main'!B1:B1");
 expect(insertFormulaReferences('Main!b2+a1','Other',{axis:'column',at:3,count:1,sheet:'main'})).toBe('Main!b2+a1');
 expect(()=>insertFormulaReferences("'Bad:Name'!A1",'Main',{axis:'row',at:1,count:1,sheet:'Main'})).toThrow();
});

test('canonical remap and matrix cases reject wrong output, false success and each missing matrix outcome',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/formula-remap.ts'),{join}=await import('node:path');const path='workflows/xlsx/formula-references.feature',source=await beforeUniformApi18Feature(path,await Bun.file(join(fixturesRoot(),path)).text()),count=(n:number)=>({implemented:n,planned:0,total:n});const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(3),cases:count(13),steps:count(42)}});
 const good=await executeAcceptance(inv(source),bindings,'formula-remap-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(13);
 const bad=await executeAcceptance(inv(source.replace('the complete replacement expression equals JSON <expected_json>','the complete replacement expression equals JSON "WRONG"')),bindings,'remap-expected-control');expect(bad.counts.cases.failed).toBe(5);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);
 const falseSuccess=bindings.map(b=>b.pattern.test('the static remapper inserts row at 1 by 1 on sheet JSON "Main"')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);const s=c.state as {error?:unknown;result?:string};if(s.error){s.error=undefined;s.result='';}}}:b);const falseReport=await executeAcceptance(inv(source),falseSuccess,'remap-refusal-control');expect(falseReport.counts.cases.failed).toBe(7);expect(falseReport.counts.steps.undefined).toBe(0);
 for(const mode of ['empty','slice','remap','span']){const corrupt=bindings.map(b=>b.pattern.test('the static analyser checks all 3 by 4 by 4 by 6 source expressions')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);const s=c.state as {matrix:Array<{slices:api.FormulaReference[][];remapped:string;wrapped:api.FormulaReference[]}>};if(mode==='empty')s.matrix=[];else if(mode==='slice')s.matrix[0]!.slices[0]=[];else if(mode==='remap')s.matrix[0]!.remapped='WRONG';else s.matrix[0]!.wrapped[0]!.start=0;}}:b);const report=await executeAcceptance(inv(source),corrupt,'remap-matrix-control');expect(report.counts.cases.failed).toBe(1);expect(report.counts.steps.failed).toBe(1);expect(report.counts.steps.undefined).toBe(0);expect(report.counts.steps.ambiguous).toBe(0);}
});
