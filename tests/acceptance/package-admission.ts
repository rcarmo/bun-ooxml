import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {admitPackage} from '../../src/opc/admission.ts';
import {OoxmlError} from '../../src/errors.ts';
import type {ZipLimits} from '../../src/opc/zip.ts';
import {admissionZip,utf8,utf16} from '../fixtures/admission.ts';
type State={bytes?:Uint8Array;before?:Uint8Array;expected?:string;error?:unknown};
const state=(c:Record<string,unknown>)=>c.state as State;
const controls=()=>{const bytes=admissionZip([['a.xml',utf8('<a/>')]]);assert.deepEqual(admitPackage(bytes).get('a.xml'),utf8('<a/>'));};
function run(s:State,limits:ZipLimits={}){assert(s.bytes);s.before=s.bytes.slice();try{admitPackage(s.bytes,limits);}catch(error){s.error=error;}}
export const bindings:StepBinding[]=[
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
