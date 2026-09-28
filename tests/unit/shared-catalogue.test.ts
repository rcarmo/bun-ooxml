import {withoutTrackingToggle} from '../helpers/execution-baseline.ts';
import {test,expect} from 'bun:test';
import {mkdtemp,mkdir,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {inventoryFeatures,executeAcceptance} from '../../scripts/gherkin.ts';import {runAcceptance} from '../../scripts/acceptance.ts';
const prefix='references/fixtures-ooxml/',catalogue=prefix+'ledgers/workflows.json',a='workflows/docx/a.feature',b='workflows/docx/b.feature';
const feature=(id:string)=>`@planned\nFeature: ${id}\n @id-${id}\n Scenario: ${id}\n  Given input ${id}\n  Then output ${id}\n`;
const ledger=()=>({schemaVersion:1,features:[a,b],workflows:[{id:'@id-a',feature:a,expandedCases:1},{id:'@id-b',feature:b,expandedCases:1}]});
const config=()=>({schemaVersion:2,catalogue,features:[{path:prefix+a,lifecycle:'implemented',runner:'bun',scenarioIds:['@id-a']}]});
async function project(){const root=await mkdtemp(join(tmpdir(),'shared-catalogue-'));await mkdir(join(root,'features'),{recursive:true});await Bun.write(join(root,prefix+a),feature('a'));await Bun.write(join(root,prefix+b),feature('b'));await Bun.write(join(root,catalogue),JSON.stringify(ledger()));await Bun.write(join(root,prefix+'manifest.json'),JSON.stringify({files:[a,b].map(path=>({path,role:'workflow'}))}));await Bun.write(join(root,'features/shared.json'),JSON.stringify(config()));return root;}
const bound=[{pattern:/^(input|output) a$/,run:()=>{}}];
test('catalogue mode includes unselected features as planned and separates shared from local obligations',async()=>{
 const root=await project();try{await Bun.write(join(root,'features/local.feature'),feature('local'));const inv=await inventoryFeatures(root);expect(inv.counts.cases).toEqual({implemented:1,planned:2,total:3});expect(inv.coverage?.shared.cases).toEqual({implemented:1,planned:1,total:2});expect(inv.coverage?.localOnly.cases).toEqual({implemented:0,planned:1,total:1});expect(inv.coverage?.catalogue).toMatchObject({path:catalogue,features:2,scenarios:2,cases:2});
 const f=inv.features.find(f=>f.path===prefix+b)!;expect(f.lifecycle).toBe('planned');expect(f.runner).toBeUndefined();expect(f.sourceSha256).toBe(new Bun.CryptoHasher('sha256').update(feature('b')).digest('hex'));
 let unsupported=0;const result=await executeAcceptance(inv,[...bound,{pattern:/^(input|output) (b|local)$/,run:()=>{unsupported++;throw Error('planned ran');}}],'catalogue');expect(result.failures).toEqual([]);expect(result.counts.cases).toEqual({passed:1,failed:0,planned:2,total:3});expect(unsupported).toBe(0);
 const report=await runAcceptance(bound,{root});expect(report.coverage).toEqual(inv.coverage);expect(report.execution.cases.planned).toBe(2);await expect(runAcceptance(bound,{root,full:true})).rejects.toThrow('planned');
 }finally{await rm(root,{recursive:true,force:true});}
});
test('canonical omission wrong ownership case-count drift duplicates and unsafe paths refuse',async()=>{
 for(const mutate of [
  (l:ReturnType<typeof ledger>)=>{l.schemaVersion=99;},
  (l:ReturnType<typeof ledger>)=>{l.features.pop();l.workflows.pop();},
  (l:ReturnType<typeof ledger>)=>{l.workflows.pop();},
  (l:ReturnType<typeof ledger>)=>{l.workflows[0]!.feature=b;},
  (l:ReturnType<typeof ledger>)=>{l.workflows[0]!.expandedCases=2;},
  (l:ReturnType<typeof ledger>)=>{l.features.push(a);},
  (l:ReturnType<typeof ledger>)=>{l.workflows.push(l.workflows[0]!);},
  (l:ReturnType<typeof ledger>)=>{l.features[0]='workflows/docx/../../escape.feature';},
 ]){const root=await project();try{const l=ledger();mutate(l);await Bun.write(join(root,catalogue),JSON.stringify(l));await expect(inventoryFeatures(root)).rejects.toThrow();}finally{await rm(root,{recursive:true,force:true});}}
});
test('catalogue mode refuses unknown activation duplicate mappings orphan files and silently implemented sources',async()=>{
 for(const mode of ['unknown','duplicate','orphan','lifecycle','missing','seal','catalogue-path','shadow','case-identity']){const root=await project();try{const c=config();
 if(mode==='unknown')c.features[0]!.path=prefix+'workflows/docx/c.feature';if(mode==='duplicate')c.features.push(c.features[0]!);if(mode==='orphan')await Bun.write(join(root,prefix+'workflows/docx/c.feature'),feature('c'));if(mode==='lifecycle')await Bun.write(join(root,prefix+b),feature('b').replace('@planned','@implemented @bun'));if(mode==='missing')await rm(join(root,prefix+b));if(mode==='seal')await Bun.write(join(root,prefix+'manifest.json'),JSON.stringify({files:[{path:a,role:'workflow'}]}));if(mode==='catalogue-path')c.catalogue=prefix+'../other.json';if(mode==='shadow')await Bun.write(join(root,'features/shadow.feature'),feature('b'));if(mode==='case-identity')await Bun.write(join(root,prefix+b),'@planned\nFeature: b\n @id-b\n Scenario Outline: Duplicate <value>\n  Given input <value>\n  Then output <value>\n  Examples:\n   | value |\n   | same |\n   | same |\n');
 await Bun.write(join(root,'features/shared.json'),JSON.stringify(c));await expect(inventoryFeatures(root)).rejects.toThrow();
 }finally{await rm(root,{recursive:true,force:true});}}
});
test('empty explicit activation retains the full planned catalogue but cannot pass acceptance',async()=>{
 const root=await project();try{await Bun.write(join(root,'features/shared.json'),JSON.stringify({...config(),features:[]}));const inv=await inventoryFeatures(root);expect(inv.counts.cases).toEqual({implemented:0,planned:2,total:2});await expect(runAcceptance([],{root})).rejects.toThrow('No implemented cases');}finally{await rm(root,{recursive:true,force:true});}
});
test('real catalogue accounts for all 791 shared cases with former local obligations centralised with only reviewed activations',async()=>{
 const inv=await inventoryFeatures(process.cwd());expect(inv.coverage?.catalogue).toMatchObject({features:62,scenarios:304,cases:791});expect(inv.coverage?.shared.cases).toEqual({implemented:730,planned:61,total:791});expect(inv.coverage?.localOnly.cases).toEqual({implemented:0,planned:0,total:0});expect(inv.counts.cases).toEqual({implemented:730,planned:61,total:791});expect(inv.counts.scenarios.total).toBe(304);
 const templates=inv.features.filter(f=>/\/template-(analysis|cache)\.feature$/.test(f.path));expect(templates).toHaveLength(2);expect(templates.flatMap(f=>f.scenarios).filter(s=>!s.scenarioId.startsWith('@id-docx-template-inventory-'))).toHaveLength(13);expect(templates.flatMap(f=>f.scenarios.map(s=>({s,f}))).filter(({s})=>!s.scenarioId.startsWith('@id-docx-template-inventory-')).every(({s,f})=>(s.lifecycle??f.lifecycle)==='planned')).toBe(true);
 const keys=inv.features.flatMap(f=>f.scenarios.filter(s=>(s.lifecycle??f.lifecycle)==='implemented').flatMap(s=>s.cases.map(c=>c.identityKey))).sort();
 const originalKeys=withoutTrackingToggle(keys);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(originalKeys)).digest('hex')).toBe('cafa5b7815a43ebe12780b1d7b9742e0e806eeea2fdc4ba36f249d7927650d01');
});
