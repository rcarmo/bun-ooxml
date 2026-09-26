import{test,expect}from'bun:test';
import{mkdtemp,rm,mkdir,cp}from'node:fs/promises';
import{tmpdir}from'node:os';
import{join}from'node:path';
import{verifyReferences}from'../../scripts/references.ts';
const root=join(import.meta.dir,'../..');
test('shared manifest, annotated tag and exact submodule commit match required pin',async()=>{expect(await verifyReferences(root)).toBe(518);});
test('missing pin refuses rather than using an implicit or generated fixture corpus',async()=>{const tmp=await mkdtemp(join(tmpdir(),'reference-pin-'));try{await expect(verifyReferences(tmp)).rejects.toThrow('Missing required shared reference pin');}finally{await rm(tmp,{recursive:true,force:true});}});
test('mutated seals and wrong tag or revision refuse in an isolated reference clone',async()=>{
 const tmp=await mkdtemp(join(tmpdir(),'reference-drift-'));try{
  await mkdir(join(tmp,'references'),{recursive:true});
  const clone=Bun.spawnSync(['git','clone','--quiet','--no-hardlinks',join(root,'references/fixtures-ooxml'),join(tmp,'references/fixtures-ooxml')]);expect(clone.exitCode).toBe(0);
  const pin=await Bun.file(join(root,'references/fixtures-ooxml.pin.json')).json();const path=join(tmp,'references/fixtures-ooxml.pin.json');
  for(const change of [{manifestSha256:'0'.repeat(64)},{sharedPackManifestSha256:'0'.repeat(64)},{commit:'0'.repeat(40)},{tag:'v99.0.0'}]){await Bun.write(path,JSON.stringify({...pin,...change}));await expect(verifyReferences(tmp)).rejects.toThrow();}
  await Bun.write(path,JSON.stringify(pin));expect(await verifyReferences(tmp)).toBe(518);
  await Bun.write(join(tmp,'references/fixtures-ooxml/fixtures/go-ooxml/testdata/default.docx'),'tampered');await expect(verifyReferences(tmp)).rejects.toThrow('Reference drift');
 }finally{await rm(tmp,{recursive:true,force:true});}
});
