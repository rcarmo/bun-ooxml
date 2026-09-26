import {test,expect} from 'bun:test';
import {runCampaign,makeInputs} from '../../scripts/property-campaign.ts';

test('seeded inputs repeat exactly and a different seed changes the corpus',()=>{
 const one=makeInputs(17,5);expect(makeInputs(17,5)).toEqual(one);expect(makeInputs(18,5)).not.toEqual(one);expect(one).toHaveLength(5);
});
test('small native campaign records actual executed families and reproducible input hashes',async()=>{
 const options={seed:17,cases:4,officeCases:2};const a=await runCampaign(options),b=await runCampaign(options);
 expect(a.status).toBe('passed');expect(a.inputSha256).toBe(b.inputSha256);
 expect(a.counts).toEqual({xmlValid:4,xmlRefused:12,zipValid:8,zipMutations:32,zipAccepted:expect.any(Number),zipRefused:expect.any(Number),docx:2,pptx:2,xlsx:2,staleRefused:4});
 expect(a.counts.zipAccepted+a.counts.zipRefused).toBe(32);expect(a.failures).toEqual([]);
});
test('campaign rejects invalid budgets and preserves failing case inputs',async()=>{
 await expect(runCampaign({seed:17,cases:0,officeCases:1})).rejects.toThrow();
 await expect(runCampaign({seed:NaN,cases:1,officeCases:1})).rejects.toThrow();
 const r=await runCampaign({seed:17,cases:2,officeCases:1,checkpoint:(id)=>{if(id==='xml/0')throw Error('injected property failure');}});
 expect(r.status).toBe('failed');expect(r.failures).toHaveLength(1);expect(r.failures[0]!.caseId).toBe('xml/0');expect(r.failures[0]!.input).toBeDefined();expect(r.failures[0]!.error).toContain('injected property failure');
});

import {timingStats} from '../../scripts/performance-campaign.ts';
test('timing reports retain raw samples and calculate median and nearest-rank p95',()=>{
 const input=[7,1,5,3,2,6,4];expect(timingStats(input)).toEqual({samplesMs:input,medianMs:4,p95Ms:7,minMs:1,maxMs:7});expect(input).toEqual([7,1,5,3,2,6,4]);expect(timingStats([2,4]).medianMs).toBe(3);for(const bad of [[],[-1],[NaN],[Infinity]])expect(()=>timingStats(bad)).toThrow();
});

import {inspectZipMutation} from '../../scripts/property-campaign.ts';
import {writePerformanceReport} from '../../scripts/performance-campaign.ts';
import {OoxmlError} from '../../src/errors.ts';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
test('mutant buffer custody is checked on accepted and typed-refusal paths',()=>{
 for(const refused of [false,true])expect(()=>inspectZipMutation(new Uint8Array([1,2,3]),bytes=>{bytes[0]=9;if(refused)throw new OoxmlError('zip-invalid','injected');return new Map();})).toThrow('caller buffer changed');
 expect(inspectZipMutation(new Uint8Array([1]),()=>{throw new OoxmlError('zip-invalid','expected');})).toEqual({refusalCode:'zip-invalid'});
 expect(()=>inspectZipMutation(new Uint8Array([1]),()=>{throw new Error('unexpected');})).toThrow('unexpected');
});
test('a failed performance run replaces old pass evidence and starts with running status',async()=>{
 const root=await mkdtemp(join(tmpdir(),'performance-report-')),path=join(root,'report.json');try{await Bun.write(path,JSON.stringify({runId:'old',status:'passed'}));await expect(writePerformanceReport(path,async()=>{expect((await Bun.file(path).json()).status).toBe('running');throw Error('injected workload failure');})).rejects.toThrow('injected workload failure');const failed=await Bun.file(path).json();expect(failed.status).toBe('failed');expect(failed.runId).not.toBe('old');expect(failed.error).toContain('injected workload failure');}finally{await rm(root,{recursive:true,force:true});}
});
