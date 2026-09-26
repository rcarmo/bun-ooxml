import {expect} from 'bun:test';
import {Document} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {trackedReplace} from '../../src/docx/redline.ts';
import {resolveRevisions,inspectRevisions} from '../../src/docx/revisions.ts';
import {inspectStories} from '../../src/docx/story.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
import {OoxmlError} from '../../src/errors.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export async function redlineFixture(kind='plain'):Promise<OpcPackage>{const d=Document.create();d.addParagraph('Payment within thirty days.');const p=await OpcPackage.open(await d.save());
 p.set('word/document.xml',p.text('word/document.xml').replace('<w:t>Payment within thirty days.</w:t>','<w:rPr><w:b/></w:rPr><w:t>Payment within </w:t></w:r><w:r><w:t>thirty days.</w:t>'));
 addPart(p,'custom/redline-sentinel.bin',new Uint8Array([5,4,3]),'application/octet-stream');
 if(kind==='ambiguous')p.set('word/document.xml',p.text('word/document.xml').replace('thirty days.','thirty days. thirty days.'));
 if(kind==='drawing')p.set('word/document.xml',p.text('word/document.xml').replace('</w:r><w:r>','</w:r><w:r><w:drawing/></w:r><w:r>'));
 if(kind==='existing')p.set('word/document.xml',p.text('word/document.xml').replace('<w:r><w:t>thirty days.</w:t></w:r>','<w:ins w:id="10" w:author="Earlier"><w:r><w:t>thirty days.</w:t></w:r></w:ins>'));
 if(kind==='protected'){addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1" w:edit="readOnly"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,'word/document.xml','http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');}
 return p;}
const options={author:'Reviewer',date:'2026-09-26T12:00:00Z'};
export const bindings:StepBinding[]=[
 {pattern:/^a native Word package with a bold fragmented payment phrase$/,run:async ctx=>{ctx.pkg=await redlineFixture();ctx.before=(ctx.pkg as OpcPackage).toBytes();}},
 {pattern:/^that phrase is replaced as tracked changes with an explicit author and date$/,run:ctx=>{ctx.result=trackedReplace(ctx.pkg as OpcPackage,'word/document.xml','within thirty','within sixty',options);}},
 {pattern:/^reopening exposes the original and current text in their review views$/,run:async ctx=>{const p=await OpcPackage.open((ctx.pkg as OpcPackage).toBytes());expect(inspectStories(p,{view:'current'}).stories[0]!.paragraphs[0]!.text).toBe('Payment within sixty days.');expect(inspectStories(p,{view:'original'}).stories[0]!.paragraphs[0]!.text).toBe('Payment within thirty days.');expect(inspectRevisions(p).revisions).toHaveLength(2);}},
 {pattern:/^accepting yields the revised text while rejecting restores the original text$/,run:async ctx=>{for(const action of ['accept','reject'] as const){const p=await OpcPackage.open((ctx.pkg as OpcPackage).toBytes());resolveRevisions(p,action);const reopened=await OpcPackage.open(p.toBytes());expect(inspectStories(reopened).stories[0]!.paragraphs[0]!.text).toBe(action==='accept'?'Payment within sixty days.':'Payment within thirty days.');ctx[action]=reopened;} }},
 {pattern:/^the starting run formatting and unrelated part bytes survive resolution$/,run:async ctx=>{const before=await OpcPackage.open(ctx.before as Uint8Array);for(const action of ['accept','reject']){const p=ctx[action] as OpcPackage;expect(elements(parseXml(p.text('word/document.xml')),'b',W).length).toBeGreaterThan(0);for(const name of before.names())if(name!=='word/document.xml')expect(p.get(name)).toEqual(before.get(name)!);}}},
 {pattern:/^a native Word tracked-edit refusal case (ambiguous|drawing|existing|protected|bad-date)$/,run:async(ctx,kind)=>{ctx.kind=kind;ctx.pkg=await redlineFixture(kind);ctx.before=(ctx.pkg as OpcPackage).toBytes();}},
 {pattern:/^the tracked replacement is attempted$/,run:ctx=>{try{trackedReplace(ctx.pkg as OpcPackage,'word/document.xml','thirty','sixty',{...options,date:ctx.kind==='bad-date'?'not-a-date':options.date});}catch(e){ctx.error=e;}}},
 {pattern:/^it refuses with a typed error and preserves the original archive bytes$/,run:ctx=>{expect(ctx.error).toBeInstanceOf(OoxmlError);expect((ctx.pkg as OpcPackage).toBytes()).toEqual(ctx.before as Uint8Array);}},
];
