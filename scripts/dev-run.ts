// Profile each test/check command; retain profiles/logs separately from scratch.
import {mkdirSync,rmSync,realpathSync} from 'node:fs';
import {join,resolve,relative} from 'node:path';
import {prepareDevPaths} from './dev-paths.ts';
const args=process.argv.slice(2),label=args.shift(),kind=args.shift();
if(!label||!/^[a-z0-9][a-z0-9-]*$/.test(label)||!['test','run','tool'].includes(kind??'')||!args.length)throw Error('Usage dev-run.ts label test|run|tool arguments');
const paths=prepareDevPaths(),root=resolve(import.meta.dir,'..');
const retained=resolve(process.env.OOXML_EVIDENCE_ROOT??join(root,'artifacts','policy-profiles'));
if(retained===paths.root||retained.startsWith(paths.root+'/'))throw Error('Evidence must be outside disposable project root');
mkdirSync(retained,{recursive:true});
if(realpathSync(retained)===realpathSync(paths.root)||realpathSync(retained).startsWith(realpathSync(paths.root)+'/'))throw Error('Evidence cannot resolve inside disposable root');
const id=`${new Date().toISOString().replace(/[:.]/g,'-')}-${crypto.randomUUID()}`,out=join(retained,label,id);mkdirSync(out,{recursive:true});
if(realpathSync(out)!==join(realpathSync(retained),relative(retained,out)))throw Error('Evidence path cannot be a symlink');
const flags=['--cpu-prof','--cpu-prof-name=cpu.cpuprofile',`--cpu-prof-dir=${out}`,'--cpu-prof-interval=1000','--heap-prof','--heap-prof-name=heap.heapprofile',`--heap-prof-dir=${out}`,'--heap-prof-interval=524288'];
process.env.OOXML_PROFILE_DIR=out;
const command=kind==='test'?[process.execPath,'test','--preload',join(root,'scripts/dev-env.ts'),'--preload',join(root,'scripts/test-profile.ts'),...args]:kind==='run'?[process.execPath,...flags,'--preload',join(root,'scripts/dev-env.ts'),...args]:args;
const child=Bun.spawn(command,{cwd:root,env:process.env,stdout:Bun.file(join(out,'stdout.log')),stderr:Bun.file(join(out,'stderr.log'))});
const exit=await child.exited;
const meta={command,exit,bun:Bun.version,revision:Bun.spawnSync(['git','rev-parse','HEAD'],{cwd:root}).stdout.toString().trim(),run:paths.run,cpuIntervalMicros:kind==='test'?Number(process.env.OOXML_CPU_INTERVAL_US??10_000):1000,heapSamplingBytes:524288,subprocesses:'Git/external oracle child processes not CPU/heap sampled by parent profiles'};
await Bun.write(join(out,'command.json'),JSON.stringify(meta,null,2)+'\n');
process.stdout.write(await Bun.file(join(out,'stdout.log')).text());process.stderr.write(await Bun.file(join(out,'stderr.log')).text());
let captureFailed=false;
if(kind!=='tool'){
 const cpu=Bun.file(join(out,'cpu.cpuprofile')),heap=Bun.file(join(out,'heap.heapprofile'));
 if(!await cpu.exists()||!await heap.exists()){captureFailed=true;await Bun.write(join(out,'analysis.json'),JSON.stringify({status:'capture-failed',exit,limits:'Missing profiles; functional exit alone does not pass profiling gate'},null,2)+'\n');}
 else{
  const p=await cpu.json(),h=await heap.json(),nodes=new Map<number,any>((p.nodes??[]).map((n:any)=>[n.id,n]));const inclusive=new Map<string,number>();
  const parents=new Map<number,number>();for(const n of p.nodes??[])for(const c of n.children??[])parents.set(c,n.id);
  for(let i=0;i<(p.samples??[]).length;i++){let id:number|undefined=p.samples[i];const seen=new Set<string>();while(id!==undefined){const n=nodes.get(id);if(!n)break;const frame=n.callFrame,key=`${frame.functionName||'(anonymous)'} ${frame.url}:${frame.lineNumber+1}`;if(!seen.has(key)){inclusive.set(key,(inclusive.get(key)??0)+(p.timeDeltas?.[i]??1000));seen.add(key)}id=parents.get(id)}}
  if(Array.isArray(p.stackTraces?.traces)){for(const trace of p.stackTraces.traces){const seen=new Set<string>();for(const frame of trace.frames??[]){const key=`${frame.name||'(anonymous)'} ${frame.sourceURL}:${frame.line??0}`;if(!seen.has(key)){inclusive.set(key,(inclusive.get(key)??0)+(p.stackTraces.interval??0.001)*1e6);seen.add(key)}}}}
  const heapRows:Array<{frame:string;bytes:number;objects:number}>=[];
  if(h.head){function visit(n:any):number{const total=(n.selfSize??0)+(n.children??[]).reduce((s:number,c:any)=>s+visit(c),0);heapRows.push({frame:`${n.callFrame?.functionName||'(anonymous)'} ${n.callFrame?.url}:${(n.callFrame?.lineNumber??-1)+1}`,bytes:total,objects:0});return total};visit(h.head)}
  else if(h.snapshot?.meta?.node_fields){const fields=h.snapshot.meta.node_fields,step=fields.length,size=fields.indexOf('self_size'),name=fields.indexOf('name');const sums=new Map<string,{bytes:number;objects:number}>();for(let i=0;i<h.nodes.length;i+=step){const key=h.strings[h.nodes[i+name]]??'(unknown)',row=sums.get(key)??{bytes:0,objects:0};row.bytes+=h.nodes[i+size];row.objects++;sums.set(key,row)};for(const [frame,row]of sums)heapRows.push({frame:frame.length>220?frame.slice(0,220)+'…':frame,...row})}
  else throw Error('Unsupported Bun heap profile shape');
  const analysis={status:'captured-needs-engineering-review',cpuSamples:p.samples?.length??p.stackTraces?.length??p.stackTraces?.traces?.length??0,cumulativeCpuTop:[...inclusive].sort((a,b)=>b[1]-a[1]).slice(0,20).map(([frame,microseconds])=>({frame,microseconds})),heapKind:h.head?'sampled-allocation-tree':'end-of-process-live-heap-snapshot',heapTop:heapRows.sort((a,b)=>b.bytes-a.bytes).slice(0,20),limits:'Bun heap snapshot is retained live memory, not total allocation churn or Go alloc_space/alloc_objects; subprocess activity excluded; inspect hotspots before acceptance'};
  if(!analysis.cpuSamples){analysis.status='empty-cpu-samples';captureFailed=true}await Bun.write(join(out,'analysis.json'),JSON.stringify(analysis,null,2)+'\n');
 }
}
console.log(`Retained command/profiles/analysis: ${out}`);
// No cleanup of other jobs, caches or retained evidence; remove only this run tmp.
if(exit===0&&!captureFailed)rmSync(paths.scratch,{recursive:true,force:true});
process.exitCode=exit||Number(captureFailed);
