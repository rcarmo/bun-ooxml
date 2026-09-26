import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {comparePackageArchives,type PackageComparison} from '../../src/opc/comparison.ts';
import {admissionZip,utf8} from '../fixtures/admission.ts';
type State={original?:Uint8Array;modified?:Uint8Array;report?:PackageComparison};
const state=(c:Record<string,unknown>)=>c.state as State;
export const bindings:StepBinding[]=[
 {pattern:/^the (original|modified) ZIP_STORED package has these ordered UTF-8 members$/,run:(c,which)=>{
  const step=c.step as {argument:{dataTable:string[][]}};const rows=step.argument.dataTable;
  assert.deepEqual(rows[0],['member','payload']);
  const entries=rows.slice(1).map(r=>{assert.equal(r.length,2);return [r[0]!,utf8(r[1]!)] as const;});
  state(c)[which as 'original'|'modified']=admissionZip(entries);
 }},
 {pattern:/^the semantic package diff compares original and modified packages$/,run:c=>{const s=state(c);assert(s.original&&s.modified);const a=s.original.slice(),b=s.modified.slice();s.report=comparePackageArchives(s.original,s.modified);assert.deepEqual(s.original,a);assert.deepEqual(s.modified,b);}},
 {pattern:/^the (equivalent_xml|changed|added|removed) member list is (.+)$/,run:(c,key,json)=>{const s=state(c);assert(s.report);assert.deepEqual(s.report[key as keyof PackageComparison],JSON.parse(json!));}},
];
