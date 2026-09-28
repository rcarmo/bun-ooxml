import assert from 'node:assert/strict';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {authorStyleSample,verifyStyleReadback,sha} from '../../scripts/xlsx-style-oracle.ts';
import {runOracleCommand} from '../../scripts/oracle-process.ts';
import {Workbook} from '../../src/xlsx/index.ts';

const roots:string[]=[];
type State={sample?:Awaited<ReturnType<typeof authorStyleSample>>;source?:Uint8Array;output?:Uint8Array;report?:any;observed?:ReturnType<typeof verifyStyleReadback>;readerExit?:number|null};
const state=(c:Record<string,unknown>)=>c.state as State;
const required=<T>(value:T|undefined):T=>{assert(value!==undefined);return value;};
export const cleanup=async()=>{await Promise.all(roots.splice(0).map(root=>rm(root,{recursive:true,force:true})));};
export const bindings:StepBinding[]=[
 {pattern:/^a saved workbook containing a newly introduced cell style$/,run:async c=>{
  const root=await mkdtemp(join(tmpdir(),'bun-xlsx-style-readback-'));roots.push(root);
  const s=state(c);s.sample=await authorStyleSample(root);
  s.source=Uint8Array.from(await Bun.file(s.sample.source).bytes());s.output=Uint8Array.from(await Bun.file(s.sample.output).bytes());
  assert.equal(s.sample.receipt.status,'committed');assert.equal(s.sample.receipt.committedChanges,1);
  assert.equal(sha(s.source),s.sample.sourceSha256);assert.equal(sha(s.output),s.sample.sha256);
  const workbook=await Workbook.open(s.output),cell=workbook.worksheet('Sheet1').getCell('A1');
  assert.equal(cell?.value,'Independent\nwrapped value');assert.equal(cell?.styleId,'1');
 }},
 {pattern:/^an independent reader opens the workbook$/,run:async c=>{
  const s=state(c),sample=required(s.sample);
  const cwd=process.cwd(),project='tests/oracles/schema/SchemaCheck.csproj';
  for(const args of [
   ['dotnet','restore',project,'--locked-mode'],
   ['dotnet','build',project,'--no-restore','--configuration','Release'],
  ]){
   const build=await runOracleCommand(args,cwd);
   assert(!build.timedOut&&!build.outputLimited,`Independent reader build exceeded command bounds: ${build.stderr}`);
   assert.equal(build.exit,0,`Independent reader build failed: ${build.stderr}`);
  }
  const validator=join(cwd,'tests/oracles/schema/bin/Release/net10.0/SchemaCheck.dll');
  assert(await Bun.file(validator).exists(),'Pinned independent reader build produced no executable');
  const result=await runOracleCommand(['dotnet',validator,sample.output],cwd);
  s.readerExit=result.exit;
  assert(!result.timedOut&&!result.outputLimited,`Independent reader exceeded command bounds: ${result.stderr}`);
  assert.equal(result.exit,0,`Independent reader failed: ${result.stderr}`);
  s.report=JSON.parse(result.stdout);
  assert.equal(sha(Uint8Array.from(await Bun.file(sample.output).bytes())),sample.sha256,'Reader changed saved output');
 }},
 {pattern:/^it reads the edited cell without an invalid style index$/,run:c=>{
  const s=state(c),sample=required(s.sample);
  assert.equal(s.readerExit,0);
  s.observed=verifyStyleReadback(required(s.report),sample.output,sample.sha256);
  assert.equal(s.observed.cellFormatCount,2);
  assert.deepEqual(s.observed.cells[0],{sheet:'Sheet1',address:'A1',value:'Independent\nwrapped value',styleIndex:1,explicitStyleIndex:true,wrapText:true,fontId:0,fillId:0,borderId:0,numberFormatId:0,baseStyleIndex:0});
 }},
 {pattern:/^its identity and version are recorded separately from the writer$/,run:async c=>{
  const s=state(c),sample=required(s.sample),report=required(s.report);
  assert.equal(report.validator,'DocumentFormat.OpenXml');assert.match(report.version,/^3\.5\.1(?:\+|$)/);assert.equal(report.profile,'Office2019');
  const writer=JSON.parse(await Bun.file(join(process.cwd(),'package.json')).text());assert.equal(writer.name,'bun-ooxml');assert.equal(typeof writer.version,'string');
  assert.notEqual(report.validator,writer.name);assert.equal(typeof Bun.version,'string');
  assert(s.observed?.cells.some((cell:{sheet:string;address:string})=>cell.sheet==='Sheet1'&&cell.address==='A1'));
  assert.equal(sha(Uint8Array.from(await Bun.file(sample.source).bytes())),sample.sourceSha256);
  assert.equal(sha(Uint8Array.from(await Bun.file(sample.output).bytes())),sample.sha256);
 }},
];
