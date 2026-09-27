import { expect, test } from 'bun:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { Document, type RunFormattingPatch } from '../../src/index.ts';
import { OpcPackage, addPart, addRelationship } from '../../src/opc/index.ts';
import { parseXml, elements, attribute } from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(pr='') {
 const d=Document.create();d.addParagraph('Alpha');const pkg=await OpcPackage.open(d.package.toBytes()),xml=pkg.text(pkg.mainPart());expect(xml.includes('<w:r>')).toBe(true);pkg.set(pkg.mainPart(),xml.replace('<w:r>','<w:r>'+pr));return Document.open(pkg.toBytes());
}

test('direct appearance reads absent fields, normalises hash colour and preserves scalar values after reopen',async()=>{
 const d=await fixture(),before=d.package.parts;expect(d.paragraphs[0]!.directRunAppearance()).toEqual([{color:null,underline:null,highlight:null,verticalAlign:null}]);
 expect(d.paragraphs[0]!.setRunFormatting({color:'#ff0000',underline:'wave',highlight:'darkBlue',verticalAlign:'superscript'})).toEqual({changedRuns:1});
 const expected={color:'ff0000',underline:'wave',highlight:'darkBlue',verticalAlign:'superscript'} as const;
 expect(d.paragraphs[0]!.directRunAppearance()).toEqual([expected]);const reopened=await Document.open(d.package.toBytes());expect(reopened.paragraphs[0]!.directRunAppearance()).toEqual([expected]);expect(reopened.paragraphs[0]!.text).toBe('Alpha');
 for(const [n,b]of before)if(n!=='word/document.xml')expect(reopened.package.get(n)).toEqual(b);
});

test('appearance edits follow schema order alongside flags and sizes and preserve unrelated property bytes',async()=>{
 const d=await fixture('<w:rPr><w:rFonts w:ascii = \'Arial\'/><w:lang w:val="en-US"/></w:rPr>');
 d.paragraphs[0]!.setRunFormatting({verticalAlign:'subscript',underline:'double',highlight:'yellow',fontSizePt:14,color:'ABCDEF',shadow:true,bold:true});
 const xml=new TextDecoder().decode(d.package.get('word/document.xml'));expect(xml).toContain('<w:rFonts w:ascii = \'Arial\'/>');expect(xml).toContain('<w:lang w:val="en-US"/>');
 expect(elements(parseXml(xml),'rPr',W)[0]!.children.map(n=>n.localName)).toEqual(['rFonts','b','shadow','color','sz','highlight','u','vertAlign','lang']);
});

test('explicit none/baseline/auto differ from null removal and same-state retains lexical spelling',async()=>{
 const d=await fixture('<w:rPr><w:color w:val = \'ff0000\'/><w:highlight w:val="none"/><w:u w:val="none"/><w:vertAlign w:val="baseline"/></w:rPr>'),p=d.paragraphs[0]!,before=d.package.toBytes();
 expect(p.setRunFormatting({color:'#ff0000',highlight:'none',underline:'none',verticalAlign:'baseline'})).toEqual({changedRuns:0});expect(d.package.toBytes()).toEqual(before);
 const values=p.directRunAppearance();values[0]!.color='changed';expect(p.directRunAppearance()[0]!.color).toBe('ff0000');
 p.setRunFormatting({color:'auto'});expect(d.paragraphs[0]!.directRunAppearance()[0]!.color).toBe('auto');expect(()=>p.directRunAppearance()).toThrow();
 d.paragraphs[0]!.setRunFormatting({color:null,highlight:null,underline:null,verticalAlign:null});expect(d.paragraphs[0]!.directRunAppearance()).toEqual([{color:null,underline:null,highlight:null,verticalAlign:null}]);
});

test('all supported highlight names and underline styles read back without applying style inheritance',async()=>{
 for(const highlight of ['black','blue','cyan','green','magenta','red','yellow','white','darkBlue','darkCyan','darkGreen','darkMagenta','darkRed','darkYellow','darkGray','lightGray','none'] as const){const d=await fixture();d.paragraphs[0]!.setRunFormatting({highlight});expect((await Document.open(d.package.toBytes())).paragraphs[0]!.directRunAppearance()[0]!.highlight).toBe(highlight);}
 for(const underline of ['none','single','double','thick','dotted','dash','wave'] as const){const d=await fixture();d.paragraphs[0]!.setRunFormatting({underline});expect(d.paragraphs[0]!.directRunAppearance()[0]!.underline).toBe(underline);}
});

test('invalid scalar values, missing values, injection and executable patches refuse atomically',async()=>{
 const d=await fixture(),before=d.package.toBytes();let called=false;
 const patches:unknown[]=[{color:'red'},{color:'FFF'},{color:'#auto'},{color:'11223344'},{color:'"/><x/>'},{underline:'wavy'},{highlight:'Yellow'},{verticalAlign:'super'},{color:123},{color:'123456',underline:true}];
 patches.push(Object.defineProperty({},'color',{enumerable:true,get(){called=true;return '123456';}}));
 for(const patch of patches){expect(()=>d.paragraphs[0]!.setRunFormatting(patch as RunFormattingPatch)).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(called).toBe(false);
});

test('selected theme colours, decorated underlines and malformed scalar leaves refuse even for removal or same-value',async()=>{
 const rows:[string,RunFormattingPatch][]=[
  ['<w:color w:val="FF0000" w:themeColor="accent1"/>',{color:'FF0000'}],
  ['<w:color w:val="FF0000" w:themeTint="80"/>',{color:null}],
  ['<w:u w:val="single" w:color="FF0000"/>',{underline:'single'}],
  ['<w:u w:val="single" w:themeColor="accent1"/>',{underline:null}],
  ['<w:highlight w:val="bogus"/>',{highlight:null}],['<w:vertAlign val="baseline"/>',{verticalAlign:'baseline'}],
  ['<w:color/>',{color:'auto'}],['<w:u/>',{underline:'single'}],['<w:vertAlign w:val="up"/>',{verticalAlign:null}],
 ];
 for(const [leaf,patch]of rows){const d=await fixture('<w:rPr>'+leaf+'</w:rPr>'),before=d.package.toBytes();expect(()=>d.paragraphs[0]!.setRunFormatting(patch)).toThrow();expect(()=>d.paragraphs[0]!.directRunAppearance()).toThrow();expect(d.package.toBytes()).toEqual(before);}
 const untouched=await fixture('<w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/><w:u w:val="single" w:color="00FF00"/></w:rPr>');
 untouched.paragraphs[0]!.setRunFormatting({bold:true});const xml=new TextDecoder().decode(untouched.package.get('word/document.xml'));expect(xml).toContain('<w:color w:val="FF0000" w:themeColor="accent1"/>');expect(xml).toContain('<w:u w:val="single" w:color="00FF00"/>');
});

test('late multi-run scalar refusal leaves all prior runs unchanged',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('</w:r>','</w:r><w:r><w:rPr><w:color w:val="auto" w:themeColor="accent1"/></w:rPr><w:t>Beta</w:t></w:r>'));
 const doc=await Document.open(pkg.toBytes()),before=doc.package.toBytes();expect(()=>doc.paragraphs[0]!.setRunFormatting({color:'123456'})).toThrow();expect(doc.package.toBytes()).toEqual(before);expect(doc.paragraphs[0]!.text).toBe('AlphaBeta');
});

test('namespace aliases and UTF-16 path saves retain scalar values and other parts',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()).replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('UTF-8','UTF-16');
 const bytes=new Uint8Array(2+text.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);
 const doc=await Document.open(pkg.toBytes());doc.paragraphs[0]!.setRunFormatting({color:'123456',underline:'dash',verticalAlign:'subscript'});const root=await mkdtemp(join(tmpdir(),'run-appearance-'));
 try{const path=join(root,'appearance.docx');await doc.save(path);const reopened=await Document.open(path);expect(reopened.paragraphs[0]!.directRunAppearance()[0]).toEqual({color:'123456',underline:'dash',highlight:null,verticalAlign:'subscript'});expect([...reopened.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(reopened.package.get(n)).toEqual(pkg.get(n));}finally{await rm(root,{recursive:true,force:true});}
});

test('protection, stale source and serialization failure preserve bytes and handle state',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');
 const locked=await Document.open(pkg.toBytes()),lb=locked.package.toBytes();expect(()=>locked.paragraphs[0]!.setRunFormatting({color:'123456'})).toThrow();expect(locked.package.toBytes()).toEqual(lb);
 const held=d.paragraphs[0]!,before=d.package.toBytes(),opc=(d as unknown as {opcPackage:OpcPackage}).opcPackage,original=opc.toBytes.bind(opc);opc.toBytes=()=>{throw Error('injected serialization');};try{expect(()=>held.setRunFormatting({highlight:'cyan'})).toThrow('injected serialization');}finally{opc.toBytes=original;}
 expect(d.package.toBytes()).toEqual(before);expect(held.directRunAppearance()[0]!.highlight).toBeNull();
 d.package.setPart('word/document.xml',new TextEncoder().encode(new TextDecoder().decode(d.package.get('word/document.xml')).replace('Alpha','Beta')));const changed=d.package.toBytes();expect(()=>held.directRunAppearance()).toThrow();expect(()=>held.setRunFormatting({color:'123456'})).toThrow();expect(d.package.toBytes()).toEqual(changed);
});

test('fifteen shared appearance getter cases execute and changed expectations fail predicates',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts');
 const {bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/run-appearance.ts');const path='workflows/docx/document-model.feature',text=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const run=(source:string)=>executeAcceptance({root:'.',features:[selectSharedScenarios(path,source,scenarioIds)],counts:{features:count(1),scenarios:count(4),cases:count(15),steps:count(46)}},bindings,'run-appearance-unit');
 const good=await run(text);expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(15);expect(good.counts.cases.planned).toBeGreaterThan(0);
 for(const [from,to,n]of [['its in-memory colour getter equals <want>','its in-memory colour getter equals wrong',3],['Underline is true and UnderlineStyle equals <style>','Underline is true and UnderlineStyle equals wrong',6],['the Highlight getter equals <colour>','the Highlight getter equals wrong',5]] as const){expect(text.includes(from)).toBe(true);const bad=await run(text.replace(from,to));expect(bad.counts.cases.failed).toBe(n);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 const corrupted=bindings.map(b=>b.pattern.test('Superscript is enabled on the first and Subscript on the second')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);(c.state as {document:Document}).document.paragraphs[1]!.setRunFormatting({verticalAlign:'baseline'});}}:b);
 const bad=await executeAcceptance({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(4),cases:count(15),steps:count(46)}},corrupted,'run-vertical-control');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);
});

test('scalar edits preserve table grid and source body text, expire handles and share staged-write rollback',async()=>{
 const d=Document.create();d.addTable(1,2);d.tables[0]!.cell(0,0).text='Cell';d.tables[0]!.cell(0,1).text='Other';const table=d.tables[0]!,p=d.paragraphs[0]!;
 p.setRunFormatting({color:'AABBCC',highlight:'cyan'});expect(table.rows).toBe(1);expect(table.columns).toBe(2);expect(table.cell(0,0).text).toBe('Cell');expect(table.cell(0,1).text).toBe('Other');expect(()=>p.directRunAppearance()).toThrow();
 const fresh=d.paragraphs[0]!,before=d.package.toBytes(),opc=(d as unknown as {opcPackage:OpcPackage}).opcPackage,set=opc.set.bind(opc);opc.set=(name,value)=>{set(name,value);throw Error('injected write');};
 try{expect(()=>fresh.setRunFormatting({underline:'dotted'})).toThrow('injected write');}finally{opc.set=set;}
 expect(d.package.toBytes()).toEqual(before);expect(fresh.directRunAppearance()[0]!.underline).toBeNull();
});
