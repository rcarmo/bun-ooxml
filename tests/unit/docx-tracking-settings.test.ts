import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {Document,OpcPackage} from '../../src/index.ts';
import {addPart,addRelationship,getContentType} from '../../src/opc/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const MIME='application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
const PART='word/options/custom.xml';
async function fixture(settings?:string){const d=Document.create();d.addParagraph('Retained text');const p=await OpcPackage.open(d.package.toBytes());addPart(p,'customXml/opaque.bin',new Uint8Array([0,255,42]),'application/octet-stream');if(settings!==undefined){addPart(p,PART,settings,MIME);addRelationship(p,p.mainPart(),REL,'options/custom.xml',{id:'settingsRef'});}return Document.open(p.toBytes());}
const source=(inner:string)=>`<w:settings xmlns:w="${W}">${inner}</w:settings>`;
const internal=(d:Document)=>(d as any).opcPackage as OpcPackage;
function unchanged(d:Document,before:ReadonlyMap<string,Uint8Array>,except:string[]=[]){for(const[n,b]of before)if(!except.includes(n))expect(d.package.get(n)).toEqual(b);}

test('tracking author toggle follows the shared getter sequence without creating revision markup',async()=>{
 const d=await fixture(),held=d.paragraphs[0]!,before=d.package.parts;
 expect(d.trackChangesEnabled).toBe(false);expect(d.trackAuthor).toBe('');
 expect(d.enableTracking('Test Author').changed).toBe(1);expect(d.trackChangesEnabled).toBe(true);expect(d.trackAuthor).toBe('Test Author');
 const enabled=d.package.toBytes();d.setTrackAuthor('New Author');expect(d.trackAuthor).toBe('New Author');expect(d.package.toBytes()).toEqual(enabled);
 expect(d.disableTracking().changed).toBe(1);expect(d.trackChangesEnabled).toBe(false);expect(d.trackAuthor).toBe('New Author');expect(held.text).toBe('Retained text');
 unchanged(d,before,['[Content_Types].xml','word/_rels/document.xml.rels']);expect(internal(d).text(internal(d).mainPart())).not.toContain('trackRevisions');
});

test('enabled setting persists through file reopen with session-only author and complete graph custody',async()=>{
 const d=await fixture(),before=d.package.parts;d.enableTracking('Reviewer & <literal>');const p=internal(d),links=p.relationships(p.mainPart()).filter(r=>r.type===REL);expect(links).toHaveLength(1);const part=links[0]!.resolved!;expect(getContentType(p,part)).toBe(MIME);
 const tree=parseXml(p.text(part)),flag=elements(tree,'trackRevisions',W);expect(tree.root.localName).toBe('settings');expect(flag).toHaveLength(1);expect(attribute(flag[0]!,'val',W)).toBe('1');expect(p.text(part)).not.toContain('Reviewer');
 const dir=await mkdtemp(join(tmpdir(),'tracking-settings-'));try{const path=join(dir,'tracked.docx');await d.save(path);const opened=await Document.open(path);expect(opened.trackChangesEnabled).toBe(true);expect(opened.trackAuthor).toBe('');expect(opened.paragraphs[0]!.text).toBe('Retained text');unchanged(opened,before,['[Content_Types].xml','word/_rels/document.xml.rels']);opened.disableTracking();await opened.save(path);expect((await Document.open(path)).trackChangesEnabled).toBe(false);}finally{await rm(dir,{recursive:true,force:true});}
});

test('custom settings path and prefix retain sibling XML bytes while insertion obeys settings order',async()=>{
 const original=`<q:settings xmlns:q="${W}" xmlns:w="urn:foreign"><q:zoom q:percent='90'/><q:revisionView q:markup='0'/><q:doNotTrackMoves/><q:compat><q:useWord2002TableStyleRules/></q:compat></q:settings>`;
 const d=await fixture(original),before=d.package.parts;d.enableTracking('Reviewer');expect(internal(d).relationships('word/document.xml').filter(r=>r.type===REL).map(r=>r.resolved)).toEqual([PART]);
 const next=internal(d).text(PART);expect(next.replace(/<w:trackRevisions[^>]*\/>/,'')).toBe(original);expect(parseXml(next).root.children.map(n=>n.localName)).toEqual(['zoom','revisionView','trackRevisions','doNotTrackMoves','compat']);unchanged(d,before,[PART]);expect((await Document.open(d.package.toBytes())).trackChangesEnabled).toBe(true);
});

test('missing false state and all supported on-off spellings preserve no-op bytes',async()=>{
 const empty=await fixture(),before=empty.package.toBytes();expect(empty.disableTracking()).toEqual({changed:0,changedParts:[]});expect(empty.package.toBytes()).toEqual(before);
 for(const[val,on]of [['',true],['1',true],['true',true],['on',true],['0',false],['false',false],['off',false]] as const){const d=await fixture(source(`<w:trackRevisions${val?` w:val='${val}'`:''}/>`)),bytes=d.package.toBytes();expect(d.trackChangesEnabled).toBe(on);expect((on?d.enableTracking('Reviewer'):d.disableTracking()).changed).toBe(0);expect(d.package.toBytes()).toEqual(bytes);}
 const self=await fixture(`<q:settings xmlns:q="${W}"/>`);self.enableTracking('Reviewer');expect((await Document.open(self.package.toBytes())).trackChangesEnabled).toBe(true);
});

test('protection, ambiguous or malformed settings refuse before package and author changes',async()=>{
 const bad=[
  '<w:documentProtection/>','<w:documentProtection w:enforcement="1"/>','<w:documentProtection enforcement="0"/>',
  '<w:trackRevisions/><w:trackRevisions/>','<w:trackRevisions val="1"/>','<w:trackRevisions w:val="yes"/>','<w:trackRevisions other="keep"/>',
  '<w:trackRevisions><w:bad/></w:trackRevisions>','<w:trackRevisions><!--retain--></w:trackRevisions>',
  '<w:compat/><w:trackRevisions/>','<w:unknown/>','<w:compat><w:trackRevisions/></w:compat>',
 ];
 for(const inner of bad){const d=await fixture(source(inner)),bytes=d.package.toBytes();d.setTrackAuthor('Previous');expect(()=>d.enableTracking('Next')).toThrow();expect(()=>d.disableTracking()).toThrow();expect(d.trackAuthor).toBe('Previous');expect(d.package.toBytes()).toEqual(bytes);}
 const unlocked=await fixture(source('<w:documentProtection w:enforcement="off"/>'));unlocked.enableTracking('Reviewer');expect(unlocked.trackChangesEnabled).toBe(true);
});

test('external duplicate orphan wrong-root wrong-type and shared settings graphs refuse unchanged',async()=>{
 for(const mode of ['external','duplicate','orphan','root','mime','shared','collision'] as const){const d=await fixture(mode==='external'||mode==='collision'?undefined:source('')),p=internal(d);
  if(mode==='external')addRelationship(p,p.mainPart(),REL,'https://example.invalid/settings',{external:true});
  if(mode==='duplicate')addRelationship(p,p.mainPart(),REL,'options/custom.xml',{id:'duplicate'});
  if(mode==='orphan'){const relPart='word/_rels/document.xml.rels';p.set(relPart,p.text(relPart).replace(/<Relationship[^>]*Id="settingsRef"[^>]*\/>/,''));}
  if(mode==='root')p.set(PART,`<w:document xmlns:w="${W}"/>`);
  if(mode==='mime')p.set('[Content_Types].xml',p.text('[Content_Types].xml').replace(MIME,'application/xml'));
  if(mode==='shared')addRelationship(p,'',REL,'word/options/custom.xml');
  if(mode==='collision')addPart(p,'word/settings.xml','opaque','application/octet-stream');
  const bytes=d.package.toBytes();expect(()=>d.enableTracking('Next')).toThrow();expect(d.trackAuthor).toBe('');expect(d.package.toBytes()).toEqual(bytes);
 }
});

test('invalid authors and stale main snapshots refuse without touching settings or invoking coercion',async()=>{
 const d=await fixture();let called=false;for(const value of ['', '  ', '\u0000', '\ud800', 1, null, {toString(){called=true;return 'Reviewer';}}]){const before=d.package.toBytes();expect(()=>d.enableTracking(value as any)).toThrow();expect(()=>d.setTrackAuthor(value as any)).toThrow();expect(d.package.toBytes()).toEqual(before);expect(d.trackAuthor).toBe('');}expect(called).toBe(false);
 const p=internal(d);p.set(p.mainPart(),p.text(p.mainPart()).replace('Retained','External'));const bytes=p.toBytes();expect(()=>d.enableTracking('Reviewer')).toThrow();expect(()=>d.disableTracking()).toThrow();expect(p.toBytes()).toEqual(bytes);
});

test('write and serialization faults restore settings graph plus the previous session author',async()=>{
 for(const existing of [false,true])for(const stage of ['set','toBytes'] as const){const d=await fixture(existing?source(''):undefined),p=internal(d),held=d.paragraphs[0]!,before=p.toBytes();d.setTrackAuthor('Previous');const original=p[stage].bind(p);
  if(stage==='set')p.set=(n,v)=>{original(n as never,v as never);throw Error('injected write');};else p.toBytes=()=>{throw Error('injected serialize');};
  try{expect(()=>d.enableTracking('Next')).toThrow('injected');}finally{if(stage==='set')p.set=original as OpcPackage['set'];else p.toBytes=original as OpcPackage['toBytes'];}
  expect(d.trackAuthor).toBe('Previous');expect(d.trackChangesEnabled).toBe(false);expect(p.toBytes()).toEqual(before);expect(held.text).toBe('Retained text');
 }
});

test('UTF16 settings retain encoding through enable disable and reopen',async()=>{
 const d=await fixture(source('<w:zoom w:percent="80"/>')),p=internal(d),text='<?xml version="1.0" encoding="UTF-16"?>'+p.text(PART),bytes=new Uint8Array(2+text.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+2*i,text.charCodeAt(i),true);p.set(PART,bytes);const before=d.package.parts;d.enableTracking('Reviewer');expect([...d.package.get(PART)!.slice(0,2)]).toEqual([255,254]);const reopened=await Document.open(d.package.toBytes());expect(reopened.trackChangesEnabled).toBe(true);reopened.disableTracking();expect((await Document.open(reopened.package.toBytes())).trackChangesEnabled).toBe(false);unchanged(reopened,before,[PART]);
});

test('canonical toggle selects one case and detects incorrect intermediate and final values',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/tracking-settings.ts');const path='workflows/docx/tracked-workflow.feature',text=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv=(source:string)=>({root:'.',features:[selectSharedScenarios(path,source,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(4)}});
 const good=await executeAcceptance(inv(text),bindings,'tracking-settings-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);
 for(const altered of [text.replace('tracking is enabled and TrackAuthor','tracking is disabled and TrackAuthor'),text.replace('TrackAuthor equals New Author','TrackAuthor equals Wrong'),text.replace('makes TrackChangesEnabled false','makes TrackChangesEnabled true')]){const bad=await executeAcceptance(inv(altered),bindings,'tracking-settings-expected');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 for(const corrupt of [(c:any)=>{c.state.document.setTrackAuthor('Wrong');},(c:any)=>{c.state.document.disableTracking();}]){const changed=bindings.map(b=>b.pattern.test('tracking is enabled with Test Author and its author is changed to New Author')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);corrupt(c);}}:b);const bad=await executeAcceptance(inv(text),changed,'tracking-settings-corrupt');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);}
});

test('the saved preference never silently converts ordinary Bun text mutations into tracked replacements',async()=>{
 const {inspectRevisions}=await import('../../src/docx/revisions.ts');const d=await fixture();d.enableTracking('Preference author');d.paragraphs[0]!.setText('Explicit plain edit');const opened=await Document.open(await d.save());expect(opened.trackChangesEnabled).toBe(true);expect(opened.paragraphs[0]!.text).toBe('Explicit plain edit');expect(inspectRevisions(internal(opened))).toEqual({revisions:[],unsupported:[]});expect(opened.trackAuthor).toBe('');
});

test('a fault after settings relationship creation rolls back all metadata and author state',async()=>{
 const d=await fixture(),p=internal(d),before=p.toBytes(),set=p.set.bind(p);d.setTrackAuthor('Previous');
 p.set=(n,v)=>{set(n,v);if(n==='word/_rels/document.xml.rels')throw Error('injected relationship write');};
 try{expect(()=>d.enableTracking('Next')).toThrow('injected relationship write');}finally{p.set=set;}
 expect(d.trackAuthor).toBe('Previous');expect(d.trackChangesEnabled).toBe(false);expect(p.toBytes()).toEqual(before);expect(p.names()).not.toContain('word/settings.xml');
});
