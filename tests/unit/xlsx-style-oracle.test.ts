import {test,expect} from 'bun:test';
import {mkdtemp,rm,lstat} from 'node:fs/promises';import {tmpdir} from 'node:os';import {join} from 'node:path';
import {authorStyleSample,verifyStyleReadback,styleCells,verifyStyleRefusal,styleFaults,corruptStyleSample,sha} from '../../scripts/xlsx-style-oracle.ts';
const hash='a'.repeat(64),path='/tmp/style-wrapped.xlsx';
function report(){return {schemaVersion:1,validator:'DocumentFormat.OpenXml',version:'3.5.1+sample',profile:'Office2019',status:'passed',results:[{file:'style-wrapped.xlsx',sha256:hash,errors:[],truncated:false,spreadsheetStyles:{cellFormatCount:2,cells:structuredClone(styleCells)}}]};}
test('independent style-reader protocol requires all saved cells, exact dependencies, text and version',()=>{
 const r=report(),before=JSON.stringify(r);expect(verifyStyleReadback(r,path,hash)).toEqual(r.results[0]!.spreadsheetStyles);expect(JSON.stringify(r)).toBe(before);
 for(const field of ['sheet','address','value','styleIndex','explicitStyleIndex','wrapText','fontId','fillId','borderId','numberFormatId','baseStyleIndex'])for(let i=0;i<3;i++){const bad=report();(bad.results[0]!.spreadsheetStyles.cells[i] as any)[field]='wrong';expect(()=>verifyStyleReadback(bad,path,hash)).toThrow();}
});
test('independent style-reader protocol rejects absent or extra cells and incomplete or stale evidence',()=>{
 for(const mutate of [(r:any)=>delete r.schemaVersion,(r:any)=>r.schemaVersion=2,(r:any)=>r.results[0].spreadsheetStyles.cells.pop(),(r:any)=>r.results[0].spreadsheetStyles.cells.push(styleCells[0]),(r:any)=>r.results[0].spreadsheetStyles.cellFormatCount=1,(r:any)=>delete r.results[0].spreadsheetStyles,(r:any)=>r.results[0].errors.push({Id:'invalid'}),(r:any)=>r.results[0].truncated=true,(r:any)=>r.results[0].exception='fail',(r:any)=>r.results[0].sha256='stale',(r:any)=>r.results[0].file='wrong.xlsx',(r:any)=>r.results.push(r.results[0]),(r:any)=>r.validator='Bun',(r:any)=>r.version='',(r:any)=>r.version='3.5.10',(r:any)=>r.profile='unknown',(r:any)=>r.status='failed']){const bad=report();mutate(bad);expect(()=>verifyStyleReadback(bad,path,hash)).toThrow();}
});
test('style refusal protocol requires specific reader failure and matching custody identity',()=>{
 const refusal=()=>({schemaVersion:1,validator:'DocumentFormat.OpenXml',version:'3.5.1+sample',profile:'Office2019',status:'failed',results:[{file:'style-wrapped.xlsx',sha256:hash,exception:'styles: cell format index 999 out of range 2'}]});
 expect(()=>verifyStyleRefusal(refusal(),path,hash,'cell format index 999')).not.toThrow();
 for(const mutate of [(r:any)=>delete r.schemaVersion,(r:any)=>r.schemaVersion=2,(r:any)=>r.status='passed',(r:any)=>r.results[0].exception='OPC failed',(r:any)=>r.results[0].exception='styles: different error',(r:any)=>r.results[0].spreadsheetStyles={},(r:any)=>r.results[0].sha256='stale',(r:any)=>r.results[0].file='other.xlsx',(r:any)=>r.results=[],(r:any)=>r.version='',(r:any)=>r.validator='bun']){const bad=refusal();mutate(bad);expect(()=>verifyStyleRefusal(bad,path,hash,'cell format index 999')).toThrow();}
});
test('style oracle fixture saves a new wrap XF with native readback and unchanged source and unrelated members',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'xlsx-style-oracle-'));try{const result=await authorStyleSample(dir);expect(result.receipt.status).toBe('committed');expect(result.sourceSha256).not.toBe(result.sha256);expect(await Bun.file(result.output).exists()).toBe(true);const bytes=await Bun.file(result.output).bytes();for(const fault of styleFaults){const bad=await corruptStyleSample(bytes,fault);expect(sha(bad)).not.toBe(result.sha256);expect(sha(bytes)).toBe(result.sha256);}await expect(corruptStyleSample(bytes,{...styleFaults[0]!,from:'not-present'})).rejects.toThrow('Control must alter');}finally{await rm(dir,{recursive:true,force:true});}await expect(lstat(dir)).rejects.toMatchObject({code:'ENOENT'});
});
