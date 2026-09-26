import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {trackedReplace} from '../../src/docx/redline.ts';
import {inspectRevisions,resolveRevisions} from '../../src/docx/revisions.ts';
import {inspectStories} from '../../src/docx/story.ts';
import {OpcPackage} from '../../src/opc/package.ts';
import {bindings,redlineFixture} from '../acceptance/redline-docx.ts';
const opts={author:'A & B',date:'2026-09-26T12:00:00Z'};
test('tracked replacement emits distinct IDs and native author/date with accept/reject algebra',async()=>{
 const p=await redlineFixture();const result=trackedReplace(p,'word/document.xml','within thirty','within sixty',opts);expect(result.revisionIds).toHaveLength(2);expect(new Set(result.revisionIds).size).toBe(2);expect(inspectRevisions(p).revisions.every(r=>r.author===opts.author&&r.date===opts.date)).toBe(true);
 const accepted=await OpcPackage.open(p.toBytes());resolveRevisions(accepted,'accept');expect(inspectStories(accepted).stories[0]!.paragraphs[0]!.text).toBe('Payment within sixty days.');
 const rejected=await OpcPackage.open(p.toBytes());resolveRevisions(rejected,'reject');expect(inspectStories(rejected).stories[0]!.paragraphs[0]!.text).toBe('Payment within thirty days.');
});
test('pure deletion and exact no-op preserve their declared outcomes',async()=>{
 const noOp=await redlineFixture(),before=noOp.toBytes();expect(trackedReplace(noOp,'word/document.xml','thirty','thirty',opts).revisionIds).toEqual([]);expect(noOp.toBytes()).toEqual(before);
 const p=await redlineFixture();expect(trackedReplace(p,'word/document.xml','thirty','',opts).revisionIds).toHaveLength(1);resolveRevisions(p,'accept');expect(inspectStories(p).stories[0]!.paragraphs[0]!.text).toBe('Payment within  days.');
});
test('invalid calendar dates and formatting-only run barriers refuse without mutation',async()=>{
 const p=await redlineFixture(),before=p.toBytes();expect(()=>trackedReplace(p,'word/document.xml','thirty','sixty',{...opts,date:'2026-02-30T12:00:00Z'})).toThrow();expect(p.toBytes()).toEqual(before);
 p.set('word/document.xml',p.text('word/document.xml').replace('</w:r><w:r>','</w:r><w:r><w:rPr><w:i/></w:rPr></w:r><w:r>'));
 const barrier=p.toBytes();expect(()=>trackedReplace(p,'word/document.xml','within thirty','within sixty',opts)).toThrow();expect(p.toBytes()).toEqual(barrier);
});

test('XML comments and literal inter-run content are barriers rather than discarded during redlining',async()=>{
 for(const gap of ['<!--keep-->','UNOWNED','<?review keep?>']){const p=await redlineFixture();p.set('word/document.xml',p.text('word/document.xml').replace('</w:r><w:r>',`</w:r>${gap}<w:r>`));const before=p.toBytes();expect(()=>trackedReplace(p,'word/document.xml','within thirty','within sixty',opts)).toThrow();expect(p.toBytes()).toEqual(before);}
});

test('text boxes and content controls refuse rather than redlining hidden unsupported stories',async()=>{
 for(const wrapper of ['txbxContent','sdt']){const p=await redlineFixture();const original=p.text('word/document.xml');p.set('word/document.xml',original.replace(/<w:p\b[^>]*>/,m=>`<w:${wrapper}>${m}`).replace('</w:p>',`</w:p></w:${wrapper}>`));const before=p.toBytes();expect(()=>trackedReplace(p,'word/document.xml','thirty','sixty',opts)).toThrow();expect(p.toBytes()).toEqual(before);}
});

test('overlapping literal occurrences refuse an ambiguous tracked target',async()=>{
 const p=await redlineFixture();p.set('word/document.xml',p.text('word/document.xml').replace('thirty days.','aaa'));const before=p.toBytes();expect(()=>trackedReplace(p,'word/document.xml','aa','x',opts)).toThrow();expect(p.toBytes()).toEqual(before);
});

test('executes all saved-outcome redline feature cases',async()=>{
 const path='workflows/native/docx-redline.feature',f=parseFeature('references/fixtures-ooxml/'+path,(await Bun.file(join(fixturesRoot(),path)).text()).replace(/^@planned/m,'@implemented @bun'));const cases=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});const inventory:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(cases.length),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inventory,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(6);
});
