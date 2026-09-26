import {test,expect} from 'bun:test';
import {parseA1Range,OoxmlError} from '../../src/index.ts';

test('direct cell and rectangle ranges preserve coordinates, absolute flags and endpoint order',()=>{
 expect(parseA1Range("'O''Brien'!$B2:C$4")).toEqual({sheet:"O'Brien",axis:'cell',first:{row:2,column:2,rowAbsolute:false,columnAbsolute:true},last:{row:4,column:3,rowAbsolute:true,columnAbsolute:false}});
 expect(parseA1Range('a3:B1')).toEqual({sheet:'',axis:'cell',first:{row:3,column:1,rowAbsolute:false,columnAbsolute:false},last:{row:1,column:2,rowAbsolute:false,columnAbsolute:false}});
 const single=parseA1Range('XFD1048576');expect(single.first).toEqual({row:1048576,column:16384,rowAbsolute:false,columnAbsolute:false});expect(single.last).toEqual(single.first);expect(single.last).not.toBe(single.first);
});
test('whole row and column references retain zero for the missing coordinate',()=>{
 expect(parseA1Range('$B:$B')).toEqual({sheet:'',axis:'column',first:{row:0,column:2,rowAbsolute:false,columnAbsolute:true},last:{row:0,column:2,rowAbsolute:false,columnAbsolute:true}});
 expect(parseA1Range('$3:$1')).toEqual({sheet:'',axis:'row',first:{row:3,column:0,rowAbsolute:true,columnAbsolute:false},last:{row:1,column:0,rowAbsolute:true,columnAbsolute:false}});
 expect(parseA1Range("'3'!1:1").sheet).toBe('3');expect(parseA1Range('Sheet!$A$1').sheet).toBe('Sheet');
});
test('quoted sheets support spaces, doubled apostrophes and Unicode without resolving a workbook',()=>{
 for(const name of ['Input Data','雪 sheet',"O'Brien",'A!B']){const source="'"+name.replaceAll("'","''")+"'!A1";expect(parseA1Range(source).sheet).toBe(name);}
 expect(parseA1Range('Sheet_1!a1').sheet).toBe('Sheet_1');
});
test('unsupported syntax refuses as one typed error without accepting partial input',()=>{
 for(const source of ['','A','1','A:B1','1:A2','A1:B','A1:2','SUM(A1)','[Book]Sheet!A1','Sheet1:Sheet2!A1','A1#','@A1','A1,B2','A1 B2','A0:A2','A1048577:A1048577','XFE:XFE','0:1','A:','1:','[X]A:B','INDIRECT(A1)','=A1','A1:B2:Z3','A01',' A1','A1 ','A1!','!A1',"'unfinished!A1","''!A1",'Sheet!Other!A1','3!A1']){
  expect(()=>parseA1Range(source)).toThrow(OoxmlError);expect(()=>parseA1Range(source)).toThrow(expect.objectContaining({code:'xlsx-range-unsupported'}));
 }
});
test('sheet qualifier policy refuses external, invalid and unsafe name forms',()=>{
 for(const source of ["'a:b'!A1","'[book]sheet'!A1","'a/b'!A1","'a\\b'!A1","'a?b'!A1","'a*b'!A1","'a\u0000b'!A1","'a\ud800b'!A1",'A B!A1'])expect(()=>parseA1Range(source)).toThrow(expect.objectContaining({code:'xlsx-range-unsupported'}));
});
test('row and column overflow, invalid argument types and excessive input refuse before work',()=>{
 for(const source of ['ZZZZZZZZZZZZZZZZZZ9999999999999999','A999999999999999999999','999999999999999999:1','$0:$1','XFE1','A1048577','a'.repeat(1024*1024+1),null,123])expect(()=>parseA1Range(source as string)).toThrow(OoxmlError);
 expect(parseA1Range('A1:XFD1048576').last).toEqual({row:1048576,column:16384,rowAbsolute:false,columnAbsolute:false});
});
test('returned coordinates are detached and no parser state leaks between calls',()=>{
 const first=parseA1Range('A1:B2');first.first.row=100;first.last.column=99;expect(parseA1Range('A1:B2').first.row).toBe(1);expect(parseA1Range('A1:B2').last.column).toBe(2);
});

test('thirteen existing direct-range canonical cases execute without activating formula analysis',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{join}=await import('node:path'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings,scenarioIds}=await import('../acceptance/xlsx-range.ts');
 const path='workflows/xlsx/formula-references.feature',f=selectSharedScenarios(path,await Bun.file(join(fixturesRoot(),path)).text(),scenarioIds),count=(n:number)=>({implemented:n,planned:0,total:n});
 const result=await executeAcceptance({root:'.',features:[f],counts:{features:count(1),scenarios:count(2),cases:count(13),steps:count(45)}},bindings,'range-unit');expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(13);expect(result.counts.cases.planned).toBe(32);
});
test('bounded UTF-8 input refuses oversize Unicode sheet qualifiers',()=>{
 const source="'"+'雪'.repeat(400000)+"'!A1";expect(source.length).toBeLessThan(1048576);expect(()=>parseA1Range(source)).toThrow(expect.objectContaining({code:'xlsx-range-unsupported'}));
});

test('boundary coordinate combinations retain both independent absolute flags',()=>{
 for(const [letters,column]of [['A',1],['Z',26],['AA',27],['XFD',16384]] as const)for(const row of [1,9,1048576])for(const ca of [false,true])for(const ra of [false,true]){
  const token=(ca?'$':'')+letters+(ra?'$':'')+row,parsed=parseA1Range(token+':'+token);expect(parsed.first).toEqual({row,column,rowAbsolute:ra,columnAbsolute:ca});expect(parsed.last).toEqual(parsed.first);
 }
});
