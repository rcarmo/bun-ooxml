import assert from 'node:assert/strict';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {Document} from '../../src/docx/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
type State={document?:Document;reopened?:Document;savedXml?:string;root?:string};
const roots:string[]=[];const state=(c:Record<string,unknown>)=>c.state as State;
export async function cleanup(){await Promise.all(roots.splice(0).map(p=>rm(p,{recursive:true,force:true})));}
export const bindings:StepBinding[]=[
 {pattern:/^a new Word document with one body paragraph and one run containing "Size sample"$/,run:c=>{const d=Document.create();d.addParagraph('Size sample');state(c).document=d;}},
 {pattern:/^that run's direct font size is set to 10.5 points$/,run:c=>{const d=state(c).document;assert(d);assert.deepEqual(d.paragraphs[0]!.setRunFormatting({fontSizePt:10.5}),{changedRuns:1});}},
 {pattern:/^the Word document is saved and reopened$/,run:async c=>{const s=state(c);if(!s.document)throw Error('Missing font-size document');s.root=await mkdtemp(join(tmpdir(),'font-size-'));roots.push(s.root);const path=join(s.root,'size.docx');await s.document.save(path);s.reopened=await Document.open(path);s.savedXml=new TextDecoder().decode(s.reopened.package.get('word/document.xml')!);}},
 {pattern:/^the paragraph text is "Size sample"$/,run:c=>{assert.equal(state(c).reopened!.paragraphs[0]!.text,'Size sample');}},
 {pattern:/^the run has exactly one direct WordprocessingML w:sz element with w:val "21"$/,run:c=>{const d=parseXml(state(c).savedXml!),runs=elements(d,'r',W);assert.equal(runs.length,1);const prs=runs[0]!.children.filter(n=>n.namespaceURI===W&&n.localName==='rPr');assert.equal(prs.length,1);const sizes=prs[0]!.children.filter(n=>n.namespaceURI===W&&n.localName==='sz');assert.equal(sizes.length,1);assert.equal(attribute(sizes[0]!,'val',W),'21');}},
 {pattern:/^the reopened run's direct font size is 10.5 points$/,run:c=>{assert.deepEqual(state(c).reopened!.paragraphs[0]!.directFontSizes(),[10.5]);}},
];
