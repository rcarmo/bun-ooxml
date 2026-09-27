import {test,expect} from 'bun:test';
import * as api from '../../src/index.ts';
const {analyzeFormulaReferences,parseA1Range}=api;
const bytes=(s:string)=>new TextEncoder().encode(s);
const slice=(s:string,r:{start:number;end:number})=>new TextDecoder().decode(bytes(s).slice(r.start,r.end));

test('static formula analysis counts references without interpreting strings, functions or exponents as cells',()=>{
 for(const[source,count]of [['IF(A1="B2",\'O\'\'Brien\'!$C$4,SUM(D1:E2))',3],['"A1"&"Sheet1!B2"',0],['LOG10(A1)+1E10',1],['=-(A1+B2)^2%',2],['TRUE',0],["'α sheet'!A1",1],['SUM(")",A1)',1],['SUM("%",A1)',1],['"="',0]] as const){const refs=analyzeFormulaReferences(source);expect(refs).toHaveLength(count);for(const ref of refs){expect(ref.start).toBeGreaterThanOrEqual(0);expect(ref.end).toBeGreaterThan(ref.start);expect(ref.end).toBeLessThanOrEqual(bytes(source).length);const{start,end,...range}=ref;expect(parseA1Range(slice(source,ref))).toEqual(range);}}
});

test('quoted sheet names and independent flags retain exact UTF8 byte spans and reversed endpoints',()=>{
 const source="'O''Brien'!$B2:C$4",refs=analyzeFormulaReferences(source);expect(refs).toHaveLength(1);expect(refs[0]).toEqual({start:0,end:bytes(source).length,sheet:"O'Brien",axis:'cell',first:{row:2,column:2,rowAbsolute:false,columnAbsolute:true},last:{row:4,column:3,rowAbsolute:true,columnAbsolute:false}});
 const unicode='"雪😀" + \'α sheet\'!$A$1 + a3:b1',r=analyzeFormulaReferences(unicode);expect(r.map(x=>slice(unicode,x))).toEqual(["'α sheet'!$A$1",'a3:b1']);expect(r[0]!.start).toBe(bytes('"雪😀" + ').length);expect(r[1]!.first.row).toBe(3);expect(r[1]!.last.row).toBe(1);
});

test('unsupported or incomplete grammar throws without returning earlier references',()=>{
 for(const source of ['INDIRECT(A1)','A1+INDIRECT(B2)','OFFSET(A1,1,1)','Name+1','XFE1','A1048577',"'unclosed!A1",'@A1','A1#','A1+SUM(2','A1 "+" B1','A1 ":" B2','A1+','A1 B1','A1,B2','SUM(,A1)','SUM(A1,)','SUM()','A1:B2:C3','Sheet1:Sheet2!A1','[Book]Sheet!A1','A:A','1:2','{1,2}','A1!','"unterminated','1E+','TRUE FALSE','UNKNOWN(A1)','(A1))','(A1',"''!A1"]){let result:unknown;expect(()=>{result=analyzeFormulaReferences(source);}).toThrow(expect.objectContaining({code:'xlsx-formula-unsupported'}));expect(result).toBeUndefined();}
});

test('bounded expression operators and nested supported functions preserve reference order',()=>{
 const source='=IF(NOT(FALSE),ROUND(ABS(-A1)^2%+MAX(B2,MIN(C3,D4))/AVERAGE(E5,F6),2),COUNT(G7,H8)&COUNTA(I9))';expect(analyzeFormulaReferences(source).map(r=>slice(source,r))).toEqual(['A1','B2','C3','D4','E5','F6','G7','H8','I9']);
 for(const op of ['+','-','*','/','^','&','=','<>','<=','>=','<','>']){const formula=`A1${op}B2`;expect(analyzeFormulaReferences(formula).map(r=>slice(formula,r))).toEqual(['A1','B2']);}
 expect(analyzeFormulaReferences('SUM("a""A1",.5,1.,1e-10,+2,-3,A1)')).toHaveLength(1);expect(analyzeFormulaReferences(' sum ( A1 , B2 ) ')).toHaveLength(2);
});

test('invalid source types and Unicode, byte, token, reference and nesting limits refuse deterministically',()=>{
 for(const source of [null,undefined,1,{},'', '   ','A1+\u0000','A1+\ud800','x'.repeat(1024*1024+1),'"'+'雪'.repeat(400000)+'"'])expect(()=>analyzeFormulaReferences(source as string)).toThrow();
 expect(()=>analyzeFormulaReferences('('.repeat(129)+'A1'+')'.repeat(129))).toThrow(expect.objectContaining({code:'xlsx-formula-limit'}));expect(()=>analyzeFormulaReferences('-'.repeat(129)+'A1')).toThrow(expect.objectContaining({code:'xlsx-formula-limit'}));expect(()=>analyzeFormulaReferences(Array(10002).fill('A1').join('+'))).toThrow(expect.objectContaining({code:'xlsx-formula-limit'}));expect(()=>analyzeFormulaReferences(Array(50002).fill('1').join('+'))).toThrow(expect.objectContaining({code:'xlsx-formula-limit'}));
 expect(analyzeFormulaReferences('('.repeat(50)+'A1'+')'.repeat(50))).toHaveLength(1);
});

test('finite expression matrix slices reparse and SUM wrapping shifts only byte positions',()=>{
 let count=0;for(const prefix of ['', 'Main!',"'Input Data'!"])for(const left of ['A1','$B2','C$3','$XFD$9'])for(const right of ['A1','$B2','C$3','$XFD$9'])for(const op of ['+','-','*','/','&','>=']){const source=prefix+left+op+prefix+right,refs=analyzeFormulaReferences(source),wrapped=analyzeFormulaReferences('SUM('+source+')');expect(refs).toHaveLength(2);expect(wrapped).toHaveLength(2);for(let i=0;i<2;i++){const r=refs[i]!;const{start,end,...value}=r;expect(parseA1Range(slice(source,r))).toEqual(value);expect(wrapped[i]).toEqual({...r,start:r.start+4,end:r.end+4});}count++;}expect(count).toBe(288);
});

test('returned references and coordinate objects are independent across calls and endpoints',()=>{
 const first=analyzeFormulaReferences('A1+A1');first[0]!.first.row=99;first[0]!.last.column=99;first.push(first[0]!);const second=analyzeFormulaReferences('A1+A1');expect(second).toHaveLength(2);expect(second[0]!.first.row).toBe(1);expect(second[0]!.last.column).toBe(1);expect(first[1]!.first.row).toBe(1);expect(second[0]!.first).not.toBe(second[0]!.last);
});

test('literal punctuation and doubled quotes never become syntax or hidden dependency records',()=>{
 for(const literal of ['(',')','+','-',':','%','=','A1','INDIRECT(A1)','a""B2']){const formula=`SUM("${literal}",A1)`;expect(analyzeFormulaReferences(formula).map(r=>slice(formula,r))).toEqual(['A1']);}
 for(const formula of ['"+"A1','A1"%"','SUM("(",A1','A1 + "closed" B2','A1 + R1C1','A1 + Table1[Column]'])expect(()=>analyzeFormulaReferences(formula)).toThrow(expect.objectContaining({code:'xlsx-formula-unsupported'}));
 const source="'😀 α'!A1 + '3'!B2 + Σheet!C3";const refs=analyzeFormulaReferences(source);expect(refs.map(r=>slice(source,r))).toEqual(["'😀 α'!A1","'3'!B2",'Σheet!C3']);expect(refs.map(r=>r.sheet)).toEqual(['😀 α','3','Σheet']);
});

test('canonical formula analysis cases detect false counts, corrupted flags/spans and fabricated refusal success',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/formula-analysis.ts'),{join}=await import('node:path');const path='workflows/xlsx/formula-references.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(4),cases:count(19),steps:count(64)}});
 const good=await executeAcceptance(inv(source),bindings,'formula-analysis-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(19);expect(good.counts.cases.planned).toBeGreaterThan(0);
 for(const[from,to]of [['it returns <count> reference records without error','it returns 99 reference records without error'],['first cell column 2 row 2','first cell column 9 row 2'],['last cell is column 3 row 4','last cell is column 3 row 9']]){const bad=await executeAcceptance(inv(source.replace(from!,to!)),bindings,'formula-expected-control');expect(bad.counts.cases.failed).toBe(from!.includes('<count>')?6:1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 for(const corruption of ['span','flags','refusal']){const altered=bindings.map(b=>b.pattern.test('the Go static formula analyser reads the source')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);const s=c.state as {refs?:api.FormulaReference[];error?:unknown};if(corruption==='refusal'&&s.error){s.error=undefined;s.refs=[];}else if(s.refs?.length){if(corruption==='span')s.refs[0]!.end=0;else if(corruption==='flags')s.refs[0]!.first.columnAbsolute=!s.refs[0]!.first.columnAbsolute;}}}:b);const bad=await executeAcceptance(inv(source),altered,'formula-corrupt-control');expect(bad.counts.cases.failed).toBe(corruption==='span'?5:corruption==='flags'?1:9);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
});

test('reference and nesting boundary values succeed without widening unsupported axis or sheet syntax',()=>{
 const refs=analyzeFormulaReferences(Array(10000).fill('$A$1').join('+'));expect(refs).toHaveLength(10000);expect(refs.at(-1)!.end).toBe(49999);expect(analyzeFormulaReferences('('.repeat(128)+'A1'+')'.repeat(128))).toHaveLength(1);
 for(const input of ['A1+XFE1','A1+B0','A1+$B:$B',"A1+'[Book]Sheet'!B2",'A1+Sheet1:Sheet2!B2','A1+SUM(A1;)','A1\t+B2','A1\n+B2'])expect(()=>analyzeFormulaReferences(input)).toThrow(expect.objectContaining({code:'xlsx-formula-unsupported'}));
});
