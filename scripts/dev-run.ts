// Ordinary development is unprofiled. Pre-release/diagnostic captures are disposed
// after analysis; retain only concise conclusions, never raw profiles or logs.
import {mkdirSync,rmSync} from 'node:fs';
import {join,resolve} from 'node:path';
import {prepareDevPaths,childPath,ownedRoot} from './dev-paths.ts';
import {analyseProfiles} from './profile-analysis.ts';
const args=process.argv.slice(2),label=args.shift(),kind=args.shift();
if(!label||!/^[a-z0-9][a-z0-9-]*$/.test(label)||!['test','run','tool'].includes(kind??'')||!args.length)throw Error('Usage dev-run.ts label test|run|tool arguments');
const mode=process.env.OOXML_PROFILE_MODE??'off';
if(!['off','pre-release','diagnostic'].includes(mode))throw Error('Invalid OOXML_PROFILE_MODE');
const profiling=mode!=='off'&&kind!=='tool';
const projectRoot=ownedRoot();process.env.PROJECT_TMP_ROOT=projectRoot;
// Never share a parent command's scratch: children/helpers get this owned run.
process.env.OOXML_RUN_DIR=childPath(projectRoot,join('runs','commands',`${label}-${process.pid}-${crypto.randomUUID()}`));
const paths=prepareDevPaths(),root=resolve(import.meta.dir,'..');
const out=join(paths.run,'capture');mkdirSync(out);
const interval=kind==='test'?Number(process.env.OOXML_CPU_INTERVAL_US??10_000):1000;
const flags=['--cpu-prof','--cpu-prof-name=cpu.cpuprofile',`--cpu-prof-dir=${out}`,`--cpu-prof-interval=${interval}`,'--heap-prof','--heap-prof-name=heap.heapprofile',`--heap-prof-dir=${out}`,'--heap-prof-interval=524288'];
const env={...process.env};delete env.OOXML_PROFILE_DIR;
if(profiling)env.OOXML_PROFILE_DIR=out;
const command=kind==='test'?[process.execPath,'test','--preload',join(root,'scripts/dev-env.ts'),...(profiling?['--preload',join(root,'scripts/test-profile.ts')]:[]),...args]:kind==='run'?[process.execPath,...(profiling?flags:[]),'--preload',join(root,'scripts/dev-env.ts'),...args]:args;
let child:ReturnType<typeof Bun.spawn>|undefined;
const abort=(signal:'SIGINT'|'SIGTERM')=>{child?.kill(signal)};
const onInt=()=>abort('SIGINT'),onTerm=()=>abort('SIGTERM');
process.on('SIGINT',onInt);process.on('SIGTERM',onTerm);
try{
 child=Bun.spawn(command,{cwd:root,env,stdout:'inherit',stderr:'inherit'});
 const exit=await child.exited;let captureFailed=false,analysis:unknown;
 if(profiling){
  try{
   const cpu=Bun.file(join(out,'cpu.cpuprofile')),heap=Bun.file(join(out,'heap.heapprofile'));
   if(!await cpu.exists()||!await heap.exists())throw Error('Missing CPU/heap capture');
   const result=analyseProfiles(await cpu.json(),await heap.json(),interval);
   captureFailed=!result.cpuSamples;analysis=result;
  }catch(error){captureFailed=true;analysis={status:'capture-failed',error:String(error),limits:'Missing/invalid capture is not passing profiling evidence'}}
  // Profiles have been consumed or diagnosed; delete immediately, including
  // failed probes. No raw archives or copying to a report directory.
  rmSync(out,{recursive:true,force:true});
  const conclusions=resolve(root,'artifacts','profile-conclusions');
  mkdirSync(conclusions,{recursive:true});
  await Bun.write(join(conclusions,`${label}-${crypto.randomUUID()}.json`),JSON.stringify({
   schemaVersion:1,label,mode,exit,bun:Bun.version,
   revision:Bun.spawnSync(['git','rev-parse','HEAD'],{cwd:root}).stdout.toString().trim(),
   analysis,rawDisposed:true
  },null,2)+'\n');
  console.log(`Profiling conclusions: ${conclusions}; raw capture disposed`);
 }
 process.exitCode=exit||Number(captureFailed);
}finally{
 // Wait for the owning child before cleanup, even if analysis fails/aborts.
 if(child){try{await child.exited}catch{}}
 process.off('SIGINT',onInt);process.off('SIGTERM',onTerm);
 rmSync(paths.run,{recursive:true,force:true});
}
