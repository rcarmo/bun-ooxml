import{test,expect}from'bun:test';
import{mkdtemp,rm,mkdir,cp,chmod,symlink,unlink}from'node:fs/promises';
import{tmpdir}from'node:os';
import{join}from'node:path';
import{verifyReferences}from'../../scripts/references.ts';
import{fixturesRoot}from'../../scripts/fixture-inputs.ts';
const root=join(import.meta.dir,'../..');
test('shared manifest, annotated tag and exact submodule commit match required pin',async()=>{expect(await verifyReferences(root)).toBe(147);});
test('missing pin refuses rather than using an implicit or generated fixture corpus',async()=>{const tmp=await mkdtemp(join(tmpdir(),'reference-pin-'));try{await expect(verifyReferences(tmp)).rejects.toThrow('Missing required shared reference pin');}finally{await rm(tmp,{recursive:true,force:true});}});
test('mutated seals and wrong tag or revision refuse in an isolated reference clone',async()=>{
 const tmp=await mkdtemp(join(tmpdir(),'reference-drift-'));try{
  await mkdir(join(tmp,'references'),{recursive:true});
  const clone=Bun.spawnSync(['git','clone','--quiet','--no-hardlinks',fixturesRoot(),join(tmp,'references/fixtures-ooxml')]);expect(clone.exitCode).toBe(0);
  const pin=await Bun.file(process.env.OOXML_REFERENCE_PIN??join(root,'references/fixtures-ooxml.pin.json')).json();const path=join(tmp,'references/fixtures-ooxml.pin.json');
  for(const change of [{manifestSha256:'0'.repeat(64)},{commit:'0'.repeat(40)},{tag:'v99.0.0'}]){await Bun.write(path,JSON.stringify({...pin,...change}));await expect(verifyReferences(tmp)).rejects.toThrow();}
  await Bun.write(path,JSON.stringify(pin));expect(await verifyReferences(tmp)).toBe(147);
  const facts=join(tmp,'references/fixtures-ooxml/facts/constants.json'),original=await Bun.file(facts).text();
  await Bun.write(facts,original+' ');await expect(verifyReferences(tmp)).rejects.toThrow('Shared reference worktree is dirty');await Bun.write(facts,original);
  const contract=join(tmp,'references/fixtures-ooxml/contracts/mutation-safety.json'),policy=await Bun.file(contract).text();
  await Bun.write(contract,policy+' ');await expect(verifyReferences(tmp)).rejects.toThrow('Reference drift');await Bun.write(contract,policy);
  const manifest=await Bun.file(join(tmp,'references/fixtures-ooxml/manifest.json')).json();
  await Bun.write(join(tmp,'references/fixtures-ooxml',manifest.files.find((f:any)=>f.role==='fixture').path),'tampered');await expect(verifyReferences(tmp)).rejects.toThrow('Reference drift');
 }finally{await rm(tmp,{recursive:true,force:true});}
});

test('index hints cannot hide modified facts or workflow ledgers from reference validation',async()=>{
 const tmp=await mkdtemp(join(tmpdir(),'reference-hidden-'));
 try{
  await mkdir(join(tmp,'references'),{recursive:true});const refs=join(tmp,'references/fixtures-ooxml');
  expect(Bun.spawnSync(['git','clone','--quiet','--no-hardlinks',fixturesRoot(),refs]).exitCode).toBe(0);
  const pin=await Bun.file(process.env.OOXML_REFERENCE_PIN??join(root,'references/fixtures-ooxml.pin.json')).json();await Bun.write(join(tmp,'references/fixtures-ooxml.pin.json'),JSON.stringify(pin));
  const git=(...args:string[])=>{const r=Bun.spawnSync(['git','-C',refs,...args]);expect(r.exitCode).toBe(0);return r.stdout.toString().trim();};
  for(const [flag,clear,path]of [['--assume-unchanged','--no-assume-unchanged','facts/constants.json'],['--skip-worktree','--no-skip-worktree','ledgers/workflows.json']]){
   const file=join(refs,path!),before=await Bun.file(file).text();git('update-index',flag!,path!);await Bun.write(file,before+' ');expect(git('status','--porcelain','--untracked-files=all')).toBe('');
   const indexBefore=await Bun.file(join(refs,'.git/index')).bytes();
   await expect(verifyReferences(tmp)).rejects.toThrow('Tracked reference');expect(await Bun.file(join(refs,'.git/index')).bytes()).toEqual(indexBefore);expect(await Bun.file(file).text()).toBe(before+' ');
   await Bun.write(file,before);git('update-index',clear!,path!);
  }
  if(process.platform!=='win32'){
   const path='facts/constants.json',file=join(refs,path),original=await Bun.file(file).bytes();
   git('config','core.filemode','false');await chmod(file,0o755);expect(git('status','--porcelain')).toBe('');
   await expect(verifyReferences(tmp)).rejects.toThrow('Tracked reference mode differs');await chmod(file,0o644);
   const target=join(tmp,'external-facts.json');await Bun.write(target,original);git('update-index','--assume-unchanged',path);await unlink(file);await symlink(target,file);expect(git('status','--porcelain')).toBe('');
   await expect(verifyReferences(tmp)).rejects.toThrow('Tracked reference must be a regular');await unlink(file);await Bun.write(file,original);git('update-index','--no-assume-unchanged',path);
  }
  expect(await verifyReferences(tmp)).toBe(147);
 }finally{await rm(tmp,{recursive:true,force:true});}
});
