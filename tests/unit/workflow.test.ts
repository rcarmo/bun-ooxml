import {afterAll,describe,expect,test} from "bun:test";
import {mkdtemp,rm,readdir,link,symlink} from "node:fs/promises";
import {tmpdir} from "node:os";
import {join} from "node:path";
import {patchOffice} from "../../src/workflow/index.ts";
import {Document,Workbook,Presentation} from "../../src/index.ts";
const fixtureRoot=join(import.meta.dir,"../../docs/contracts/shared-v2/pack/fixtures");
const roots:string[]=[];
afterAll(async()=>{await Promise.all(roots.map(p=>rm(p,{recursive:true,force:true})));});
async function setup(name:string){const root=await mkdtemp(join(tmpdir(),"bun-patch-"));roots.push(root);const source=join(root,name);await Bun.write(source,Bun.file(join(fixtureRoot,name)));return {root,source,output:join(root,'output.'+name.split('.').at(-1))};}
const hash=(b:Uint8Array)=>new Bun.CryptoHasher('sha256').update(b).digest('hex');
describe('native staged workflow',()=>{
 test('previews report matched targets but never write source or an existing destination',async()=>{
  const {root,source,output}=await setup('present-placeholder.docx');await Bun.write(output,'existing');const before=await Bun.file(source).bytes();
  const r=await patchOffice({source,output,mode:'dry_run',changes:[{target:'<Present>',value:'changed'}]});
  expect(r.status).toBe('preview');expect(r.committedChanges).toBe(0);expect(r.results[0]).toMatchObject({target:'<Present>',value:'changed',matched:1,status:'matched'});
  expect(await Bun.file(source).bytes()).toEqual(before);expect(await Bun.file(output).text()).toBe('existing');expect((await readdir(root)).sort()).toEqual([source.split('/').at(-1)!,output.split('/').at(-1)!].sort());
 });
 test('strict batches reject a later missing target before saving an earlier match',async()=>{
  const {source,output}=await setup('default-style.xlsx');await Bun.write(output,'existing');const before=await Bun.file(source).bytes();
  const r=await patchOffice({source,output,mode:'strict',changes:[{target:'A1',value:'after'},{target:'Missing!B1',value:123}]});
  expect(r.status).toBe('refused');expect(r.committedChanges).toBe(0);expect(r.results[1]).toMatchObject({target:'Missing!B1',matched:0,status:'unmatched'});
  expect(await Bun.file(source).bytes()).toEqual(before);expect(await Bun.file(output).text()).toBe('existing');
 });
 test('distinct output accumulates two PPTX edits and counts actual commits',async()=>{
  const {source,output}=await setup('title-and-subtitle.pptx');const before=await Bun.file(source).bytes();
  const r=await patchOffice({source,output,mode:'safe',changes:[{target:'slide:1/title',value:'New title'},{target:'slide:1/subtitle',value:'New subtitle'}]});
  expect(r.status).toBe('committed');expect(r.committedChanges).toBe(2);expect(r.results.every(x=>x.status==='committed')).toBe(true);
  expect((await Presentation.open(output)).slides[0]!.inspectText('').map(x=>x.text)).toEqual(['New title','New subtitle']);expect(await Bun.file(source).bytes()).toEqual(before);
 });
 test('source fingerprints and safe path aliases refuse without writes',async()=>{
  const {source,output}=await setup('present-placeholder.docx');const before=await Bun.file(source).bytes();
  const changes=[{target:'<Present>',value:'changed'}];
  expect((await patchOffice({source,mode:'strict',changes,expectedSourceSha256:'0'.repeat(64)})).error?.code).toBe('workflow-stale-source');
  expect((await patchOffice({source,output:source,mode:'safe',changes})).error?.code).toBe('workflow-output-alias');
  await link(source,output);expect((await patchOffice({source,output,mode:'safe',changes})).error?.code).toBe('workflow-output-alias');
  expect(await Bun.file(source).bytes()).toEqual(before);
 });
 test('destination fingerprint and symlink protections leave existing bytes untouched',async()=>{
  const {source,output,root}=await setup('present-placeholder.docx');await Bun.write(output,'old');const changes=[{target:'<Present>',value:'new'}];
  expect((await patchOffice({source,output,mode:'safe',changes,expectedDestinationSha256:null})).error?.code).toBe('workflow-stale-destination');
  const alias=join(root,'alias.docx');await symlink(output,alias);expect((await patchOffice({source,output:alias,mode:'safe',changes})).error?.code).toBe('workflow-symlink');expect(await Bun.file(output).text()).toBe('old');
 });
 test('XLSX style closure and cache receipts reflect staged output, not guessed input count',async()=>{
  const first=await setup('default-style.xlsx');const r=await patchOffice({source:first.source,output:first.output,mode:'safe',multilineWrap:true,changes:[{target:'A1',value:'first\nsecond'}]});
  expect(r.status).toBe('committed');expect(r.committedChanges).toBe(1);expect(r.changedParts).toContain('xl/styles.xml');expect((await Workbook.open(first.output)).worksheet('Sheet').getCell('A1')?.value).toBe('first\nsecond');
  const second=await setup('cross-sheet-cache.xlsx');const cache=await patchOffice({source:second.source,output:second.output,mode:'safe',calculationPolicy:'invalidate-without-recalculation',changes:[{target:'Input!A1',value:10}]});
  expect(cache.status).toBe('committed');expect(cache.calculationState).toBe('recalculation-required');expect((await Workbook.open(second.output)).worksheet('Calc').getCell('A1')).toMatchObject({kind:'formula',formula:'Input!A1*2',cached:null});
 });
 test('invalid XML replacements refuse asynchronously without touching either file',async()=>{
  const {source,output}=await setup('present-placeholder.docx');await Bun.write(output,'keep');const before=await Bun.file(source).bytes();
  const r=await patchOffice({source,output,mode:'safe',changes:[{target:'<Present>',value:'bad\u0001text'}]});
  expect(r.status).toBe('refused');expect(r.committedChanges).toBe(0);expect(await Bun.file(source).bytes()).toEqual(before);expect(await Bun.file(output).text()).toBe('keep');
 });
 test('concurrent guarded writers serialize and only one source revision can commit',async()=>{
  const {source}=await setup('present-placeholder.docx');const expectedSourceSha256=hash(await Bun.file(source).bytes());
  const results=await Promise.all(['one','two'].map(value=>patchOffice({source,mode:'strict',expectedSourceSha256,changes:[{target:'<Present>',value}]})));
  expect(results.map(r=>r.status).sort()).toEqual(['committed','refused']);expect(results.find(r=>r.status==='refused')?.error?.code).toBe('workflow-stale-source');
 });
 test('DOCX overlapping text ranges refuse even when neither target contains the other',async()=>{
  const {source,output}=await setup('present-placeholder.docx');const before=await Bun.file(source).bytes();
  const r=await patchOffice({source,output,mode:'safe',changes:[{target:'<Pre',value:'x'},{target:'resent>',value:'y'}]});
  expect(r.status).toBe('refused');expect(r.error?.code).toBe('workflow-overlapping-targets');expect(await Bun.file(output).exists()).toBe(false);expect(await Bun.file(source).bytes()).toEqual(before);
 });
 test('no-op and duplicate targets cannot inflate committed count',async()=>{
  const {source,output}=await setup('present-placeholder.docx');const before=await Bun.file(source).bytes();const changes=[{target:'<Present>',value:'<Present>'}];
  const r=await patchOffice({source,output,mode:'safe',changes});expect(r.status).toBe('committed');expect(r.committedChanges).toBe(0);expect(r.results[0]?.status).toBe('unchanged');expect(hash(await Bun.file(output).bytes())).toBe(hash(before));
  const dup=await patchOffice({source,output,mode:'safe',changes:[...changes,...changes]});expect(dup.status).toBe('refused');expect(dup.error?.code).toBe('workflow-overlapping-targets');
 });
});
