import {test,expect} from 'bun:test';
import {runOracleCommand} from '../../scripts/oracle-process.ts';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';

test('oracle runner captures success and rejects missing commands',async()=>{
 const r=await runOracleCommand([process.execPath,'-e','console.log("out");console.error("err")'],process.cwd(),5000);
 expect(r.exit).toBe(0);expect(r.timedOut).toBe(false);expect(r.stdout.trim()).toBe('out');expect(r.stderr.trim()).toBe('err');
 await expect(runOracleCommand(['/no/such/oracle-command'],process.cwd(),100)).rejects.toThrow();
});
test('oracle timeout kills descendants retaining output pipes and bounds completion',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'oracle-process-test-')),marker=join(dir,'survived');
 try{
  const child=`console.log("child-ready");setTimeout(()=>Bun.write(${JSON.stringify(marker)},"survived"),2200);setInterval(()=>{},1000)`;
  const parent=`Bun.spawn([process.execPath,"-e",${JSON.stringify(child)}],{stdout:"inherit",stderr:"inherit"});setInterval(()=>{},1000)`;
  const r=await runOracleCommand([process.execPath,'-e',parent],process.cwd(),500);expect(r.stdout).toContain('child-ready');expect(r.timedOut).toBe(true);expect(r.ms).toBeLessThan(2000);await Bun.sleep(2400);expect(await Bun.file(marker).exists()).toBe(false);
 }finally{await rm(dir,{recursive:true,force:true});}
});
test('oracle output is bounded and excessive output terminates the command group',async()=>{
 const r=await runOracleCommand([process.execPath,'-e','setInterval(()=>console.log("x".repeat(4096)),1)'],process.cwd(),5000,4096);
 expect(r.outputLimited).toBe(true);expect(Buffer.byteLength(r.stdout+r.stderr)).toBeLessThanOrEqual(4096);expect(r.ms).toBeLessThan(2000);
});
