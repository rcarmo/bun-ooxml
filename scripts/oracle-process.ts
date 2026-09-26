import {spawn} from 'node:child_process';

export interface OracleCommandResult {
  args:string[];exit:number|null;signal:string|null;timedOut:boolean;outputLimited:boolean;
  ms:number;stdout:string;stderr:string;
}

/** POSIX development runner. One process group per command, bounded captured output
 * and a bounded post-kill drain even when a descendant escapes the group. */
export function runOracleCommand(args:string[],cwd:string,timeoutMs=90_000,maxOutputBytes=4*1024*1024):Promise<OracleCommandResult>{
 if(process.platform==='win32')return Promise.reject(Error('Office oracle runner requires POSIX process groups'));
 if(!args.length||!Number.isSafeInteger(timeoutMs)||timeoutMs<1||!Number.isSafeInteger(maxOutputBytes)||maxOutputBytes<1)return Promise.reject(Error('Invalid oracle command limits'));
 return new Promise((resolve,reject)=>{
  const start=Date.now(),child=spawn(args[0]!,args.slice(1),{cwd,detached:true,stdio:['ignore','pipe','pipe']});
  const stdout:Buffer[]=[],stderr:Buffer[]=[];let bytes=0,done=false,timedOut=false,outputLimited=false,exit:number|null=null,signal:string|null=null;
  let drain:ReturnType<typeof setTimeout>|undefined;
  const killGroup=()=>{if(child.pid){try{process.kill(-child.pid,'SIGKILL');}catch(error){if((error as NodeJS.ErrnoException).code!=='ESRCH')child.kill('SIGKILL');}}};
  const finish=()=>{
   if(done)return;done=true;clearTimeout(deadline);if(drain)clearTimeout(drain);
   killGroup();child.stdout.destroy();child.stderr.destroy();child.unref();
   resolve({args:[...args],exit,signal,timedOut,outputLimited,ms:Date.now()-start,stdout:Buffer.concat(stdout).toString(),stderr:Buffer.concat(stderr).toString()});
  };
  const stop=()=>{killGroup();if(!drain)drain=setTimeout(finish,1000);};
  const deadline=setTimeout(()=>{timedOut=true;stop();},timeoutMs);
  const collect=(chunks:Buffer[],chunk:Buffer)=>{const remaining=Math.max(0,maxOutputBytes-bytes);if(remaining)chunks.push(chunk.subarray(0,remaining));bytes+=chunk.length;if(bytes>maxOutputBytes){outputLimited=true;stop();}};
  child.stdout.on('data',(b:Buffer)=>collect(stdout,b));child.stderr.on('data',(b:Buffer)=>collect(stderr,b));
  child.on('exit',(code,sig)=>{exit=code;signal=sig;});
  child.on('close',finish);
  child.on('error',error=>{if(done)return;done=true;clearTimeout(deadline);if(drain)clearTimeout(drain);killGroup();child.stdout.destroy();child.stderr.destroy();reject(error);});
 });
}
