import {test,expect} from 'bun:test';
import {mkdtempSync,rmSync,readdirSync,existsSync} from 'node:fs';
import {join,resolve} from 'node:path';
import {tmpdir} from 'node:os';
import {analyseProfiles} from '../../scripts/profile-analysis.ts';
const root=resolve(import.meta.dir,'../..');
test('profile conclusions stay bounded and report empty/invalid captures',()=>{
 const cpu={nodes:[{id:1,callFrame:{functionName:'work',url:'file.ts',lineNumber:0}}],samples:[1,1],timeDeltas:[1000,2000]};
 const heap={snapshot:{meta:{node_fields:['name','self_size']}},nodes:[0,32,1,64],strings:['Object','x'.repeat(400)]};
 const result=analyseProfiles(cpu,heap,1000);expect(result.cpuSamples).toBe(2);expect(result.cumulativeCpuTop[0]?.microseconds).toBe(3000);expect(result.totalHeapBytes).toBe(96);expect(result.heapTop[0]!.frame.length).toBeLessThan(200);
 expect(analyseProfiles({...cpu,samples:[]},heap,1000).status).toBe('empty-cpu-samples');
 expect(()=>analyseProfiles(cpu,{},1000)).toThrow('Unsupported');
});
async function command(mode:string,label:string,body:string|string[]){
 const dir=mkdtempSync(join(tmpdir(),'lifecycle-')),base=join(dir,'base'),project=join(base,'bun-ooxml');
 const conclusions=join(root,'artifacts/profile-conclusions');
 try{
  const env={...process.env,PROJECT_TMP_BASE:base,PROJECT_TMP_ROOT:project,OOXML_PROFILE_MODE:mode};
  const child=Bun.spawn([process.execPath,join(root,'scripts/dev-run.ts'),label,'run',...(typeof body==='string'?['-e',body]:body)],{cwd:root,env,stdout:'pipe',stderr:'pipe'});
  const [exit,stdout,stderr]=await Promise.all([child.exited,new Response(child.stdout).text(),new Response(child.stderr).text()]);
  const summaries=existsSync(conclusions)?readdirSync(conclusions).filter(f=>f.startsWith(label+'-')):[];
  expect(existsSync(join(project,'runs/commands'))?readdirSync(join(project,'runs/commands')):[]).toEqual([]);
  return {exit,stdout,stderr,summaries:await Promise.all(summaries.map(async f=>(await Bun.file(join(conclusions,f)).json())))};
 }finally{
  if(existsSync(conclusions))for(const f of readdirSync(conclusions))if(f.startsWith(label+'-'))rmSync(join(conclusions,f));
  rmSync(dir,{recursive:true,force:true});
 }
}
test('ordinary development does not capture or archive profiles',async()=>{
 const r=await command('off','lifecycle-off',`console.log(process.env.OOXML_PROFILE_DIR??'no-profile')`);
 expect(r.exit).toBe(0);expect(r.stdout).toContain('no-profile');expect(r.summaries).toEqual([]);
});
test('diagnostic CPU/heap is analysed then disposed with concise conclusions',async()=>{
 const r=await command('diagnostic','lifecycle-diagnostic','const end=Date.now()+250;let total=0;while(Date.now()<end)total++;console.log(total)');
 expect(r.exit).toBe(0);expect(r.summaries).toHaveLength(1);expect(r.summaries[0].rawDisposed).toBe(true);expect(r.summaries[0].analysis.cpuSamples).toBeGreaterThan(0);
},15_000);
test('failed command profiles are analysed and disposed rather than archived',async()=>{
 const r=await command('diagnostic','lifecycle-failure','process.exit(7)');
 expect(r.exit).toBe(7);expect(r.summaries).toHaveLength(1);expect(r.summaries[0].rawDisposed).toBe(true);expect(r.summaries[0].exit).toBe(7);
},15_000);
test('missing capture is explicitly failed and scratch is disposed',async()=>{
 const r=await command('diagnostic','lifecycle-missing',['--invalid-lifecycle-option']);
 expect(r.exit).not.toBe(0);expect(r.summaries).toHaveLength(1);expect(r.summaries[0].rawDisposed).toBe(true);expect(r.summaries[0].analysis.status).toBe('capture-failed');
},15_000);
