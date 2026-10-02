import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {admitPackage,admitPackageFile} from '../../src/opc/admission.ts';
import {spyOn} from 'bun:test';
import {OoxmlError} from '../../src/errors.ts';
import type {ZipLimits} from '../../src/opc/zip.ts';
import {admissionZip,utf8,utf16} from '../fixtures/admission.ts';
import {fixturePath} from '../../scripts/fixture-inputs.ts';
import {readFileSync} from 'node:fs';
type State={bytes?:Uint8Array;before?:Uint8Array;expected?:string;error?:unknown;delivered?:Map<string,Uint8Array>;preflight?:boolean;budget?:'maxArchiveBytes'|'maxEntries'};
const state=(c:Record<string,unknown>)=>c.state as State;
const controls=()=>{const bytes=admissionZip([['a.xml',utf8('<a/>')]]);assert.deepEqual(admitPackage(bytes).get('a.xml'),utf8('<a/>'));};
function run(s:State,limits:ZipLimits={}){assert(s.bytes);s.before=s.bytes.slice();try{s.delivered=admitPackage(s.bytes,limits);}catch(error){s.error=error;}}
export const bindings:StepBinding[]=[
 {pattern:/^the byte-sealed valid DOCX archive (fixture-[a-f0-9]{64}) and a separate caller byte snapshot$/,run:(c,id)=>{
  assert.equal(id,'fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5');
  const s=state(c);s.bytes=Uint8Array.from(readFileSync(fixturePath(id!)));s.before=s.bytes.slice();
  assert(admitPackage(s.bytes).size>0,'sealed input must admit with default limits');assert.deepEqual(s.bytes,s.before);
 }},
 {pattern:/^only the (source bytes|entry count) admission budget is set to -1$/,run:(c,budget)=>{const s=state(c);s.budget=budget==='source bytes'?'maxArchiveBytes':'maxEntries';}},
 {pattern:/^bounded package admission checks that archive$/,run:c=>{const s=state(c);assert(s.budget);assert(s.bytes);run(s,{[s.budget]:-1});}},
 {pattern:/^it refuses the invalid caller budget before reading source metadata or ZIP members and returns no package or parts$/,run:async c=>{
  const s=state(c);assert(s.bytes);assert(s.budget);assert.equal(s.delivered,undefined);assert(s.error instanceof OoxmlError);assert.equal(s.error.code,'package-admission-limit-invalid');
  const source=fixturePath('fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5'),reads:string[]=[];
  const probe=spyOn(Bun,'file').mockImplementation(()=>{reads.push('source-file');throw Error('source read before argument refusal');});
  try{await assert.rejects(()=>admitPackageFile(source,{[s.budget!]:-1}),{code:'package-admission-limit-invalid'});assert.deepEqual(reads,[]);}
  finally{probe.mockRestore();}
  const admitted=await admitPackageFile(source);assert(admitted.has('word/document.xml'));
  const broken=Uint8Array.of(0x00);assert.throws(()=>admitPackage(broken,{[s.budget!]:-1}),{code:'package-admission-limit-invalid'});s.preflight=true;
 }},
 {pattern:/^the refusal is an invalid-argument result, not a resource-limit or malformed-archive result$/,run:c=>{
  const s=state(c);assert(s.preflight);assert(s.error instanceof OoxmlError);assert.equal(s.error.code,'package-admission-limit-invalid');
  assert.notEqual(s.error.code,'zip-too-many-entries');assert.notEqual(s.error.code,'zip-archive-too-large');assert.notEqual(s.error.code,'zip-end-record-missing');
 }},
 {pattern:/^the caller's archive bytes remain unchanged$/,run:c=>{const s=state(c);assert(s.error);assert.deepEqual(s.bytes,s.before);}},
 {pattern:/^an ordered ZIP_STORED archive has member pairs encoded as JSON (.+)$/,run:(c,json)=>{
  const entries=JSON.parse(json!) as Array<[string,string]>,s=state(c);s.bytes=admissionZip(entries.map(([name,text])=>[name,utf8(text)]));
  s.expected=entries.length>1?'zip-duplicate-entry':entries[0]![0].endsWith('/')?'zip-directory-entry-invalid':'zip-name-invalid';
 }},
 {pattern:/^a ZIP_DEFLATED archive contains a.xml with UTF-8 XML enclosing exactly 10000 spaces between <a> and <\/a>$/,run:c=>{state(c).bytes=admissionZip([['a.xml',utf8('<a>'+' '.repeat(10000)+'</a>')]],8);}},
 {pattern:/^a ZIP_BZIP2 archive contains a.xml with UTF-8 text <a\/>$/,run:c=>{const s=state(c);s.bytes=admissionZip([['a.xml',utf8('<a/>')]],12);s.expected='zip-method-unsupported';}},
 {pattern:/^a ZIP_STORED archive contains a.xml with (UTF-8|UTF-16 with BOM) text (.+)$/,run:(c,encoding,text)=>{
  const s=state(c);s.bytes=admissionZip([['a.xml',encoding==='UTF-8'?utf8(text!):utf16(text!)]]);s.expected=text!.startsWith('<!DOCTYPE')?'XML_DTD_FORBIDDEN':'XML_MALFORMED';
 }},
 {pattern:/^the package admission guard checks the archive with default limits$/,run:c=>run(state(c))},
 {pattern:/^the package admission guard checks the archive with only (max_members|max_member_bytes|max_total_bytes|max_ratio) set to (\d+)$/,run:(c,limit,value)=>{
  const s=state(c),policy={max_members:['maxEntries','zip-too-many-entries'],max_member_bytes:['maxEntryBytes','zip-entry-too-large'],max_total_bytes:['maxTotalBytes','zip-total-too-large'],max_ratio:['maxCompressionRatio','zip-compression-ratio-exceeded']}[limit!]!;
  assert(policy);s.expected=policy[1];assert(s.bytes);assert.equal(admitPackage(s.bytes).get('a.xml')!.length,10007);run(s,{[policy[0]!]:Number(value)});
 }},
 {pattern:/^package admission is refused$/,run:c=>{const s=state(c);assert(s.error instanceof OoxmlError,'expected typed admission refusal');assert.equal(s.error.code,s.expected);assert.deepEqual(s.bytes,s.before,'refusal must not alter caller bytes');controls();}},
 {pattern:/^the admission error contains compression$/,run:c=>{const error=state(c).error;assert(error instanceof OoxmlError);assert.match(error.message,/compression/);}},
];
