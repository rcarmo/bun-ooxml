// Bun 1.4.2 test CLI does not flush CLI profiler flags. Profile in process.
import {afterAll} from 'bun:test';
import {Session} from 'node:inspector';
import {writeFileSync} from 'node:fs';
import {join} from 'node:path';
const output=process.env.OOXML_PROFILE_DIR;
if(!output)throw Error('Tests require dev-run.ts profiling entrypoint');
const interval=Number(process.env.OOXML_CPU_INTERVAL_US??10_000);
if(!Number.isSafeInteger(interval)||interval<1000||interval>100_000)throw Error('Invalid CPU sampling interval');
const session=new Session();session.connect();
function post(method:string,params?:Record<string,unknown>):Promise<any>{return new Promise((resolve,reject)=>session.post(method,params??{},(error,result)=>error?reject(error):resolve(result)))}
await post('Profiler.enable');await post('Profiler.setSamplingInterval',{interval});await post('Profiler.start');
afterAll(async()=>{
 const {profile}=await post('Profiler.stop');
 writeFileSync(join(output,'cpu.cpuprofile'),JSON.stringify(profile));
 writeFileSync(join(output,'heap.heapprofile'),Bun.generateHeapSnapshot('v8'));
 session.disconnect();
},120_000);
