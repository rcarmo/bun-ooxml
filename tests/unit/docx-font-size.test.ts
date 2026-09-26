import {test,expect} from 'bun:test';
import {Document,OpcPackage} from '../../src/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',part='word/document.xml';
const xml=(d:Document)=>new TextDecoder().decode(d.package.get(part)!);
function formattingPackage():Uint8Array {
 const d=Document.create();d.package.setPart(part,new TextEncoder().encode(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:color w:val="008800"/></w:rPr><w:t>Alpha</w:t></w:r><w:r><w:t xml:space="preserve"> Beta</w:t></w:r><w:r><w:rPr/><w:t xml:space="preserve"> Gamma</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`));return d.package.toBytes();
}
test('direct font size writes exactly one half-point value and reopens without changing text or other parts',async()=>{
 const d=Document.create();d.addParagraph('Size sample');const source=d.package.toBytes(),before=await OpcPackage.open(source),opened=await Document.open(source);
 expect(opened.paragraphs[0]!.directFontSizes()).toEqual([null]);expect(opened.paragraphs[0]!.setRunFormatting({fontSizePt:10.5})).toEqual({changedRuns:1});
 const saved=opened.package.toBytes(),reopened=await Document.open(saved),nodes=elements(parseXml(xml(reopened)),'sz',W);
 expect(nodes).toHaveLength(1);expect(attribute(nodes[0]!,'val',W)).toBe('21');expect(reopened.paragraphs[0]!.directFontSizes()).toEqual([10.5]);expect(reopened.paragraphs[0]!.text).toBe('Size sample');
 const after=await OpcPackage.open(saved);for(const name of before.names())if(name!==part)expect(after.get(name)).toEqual(before.get(name));expect(after.names()).toEqual(before.names());
});
test('font-size set, remove and combined patches preserve schema ordering and unrelated direct properties',async()=>{
 const d=await Document.open(formattingPackage()),original=xml(d);expect(d.paragraphs[0]!.setRunFormatting({bold:false,italic:true,fontSizePt:12}).changedRuns).toBe(3);
 const runs=elements(parseXml(xml(d)),'r',W);for(const r of runs){const names=r.children.find(n=>n.localName==='rPr')!.children.map(n=>n.localName);expect(names.indexOf('b')).toBeLessThan(names.indexOf('i'));expect(names.indexOf('i')).toBeLessThan(names.indexOf('sz'));}
 expect(xml(d)).toContain('<w:color w:val="008800"/>');expect(d.paragraphs[0]!.directFontSizes()).toEqual([12,12,12]);
 expect(d.paragraphs[0]!.setRunFormatting({fontSizePt:null}).changedRuns).toBe(3);expect(d.paragraphs[0]!.directFontSizes()).toEqual([null,null,null]);expect(d.paragraphs[0]!.text).toBe('Alpha Beta Gamma');expect(original).toContain('Alpha');
});
test('equal size including leading-zero halfpoints is byte-identical and keeps handles live',async()=>{
 const p=await OpcPackage.open(formattingPackage());p.set(part,p.text(part).replace('<w:color w:val="008800"/>','<w:color w:val="008800"/><w:sz w:val="021"/>'));
 const d=await Document.open(p.toBytes()),h=d.paragraphs[0]!;h.setRunFormatting({fontSizePt:10.5});const current=d.paragraphs[0]!,bytes=d.package.toBytes();expect(current.setRunFormatting({fontSizePt:10.5})).toEqual({changedRuns:0});expect(d.package.toBytes()).toEqual(bytes);expect(current.text).toBe('Alpha Beta Gamma');
});
test('font-size validation and ambiguous existing sizes refuse atomically',async()=>{
 const d=await Document.open(formattingPackage()),before=d.package.toBytes(),h=d.paragraphs[0]!;
 for(const size of [NaN,Infinity,-1,0,10.25,'12']){expect(()=>h.setRunFormatting({fontSizePt:size as number})).toThrow();expect(d.package.toBytes()).toEqual(before);expect(h.text).toBe('Alpha Beta Gamma');}
 for(const property of ['<w:sz/>','<w:sz val="21"/>','<w:sz w:val="bad"/>','<w:sz w:val="21"/><w:sz w:val="22"/>']){
  const p=await OpcPackage.open(formattingPackage());p.set(part,p.text(part).replace('<w:color w:val="008800"/>','<w:color w:val="008800"/>'+property));const broken=await Document.open(p.toBytes()),bytes=broken.package.toBytes();
  expect(()=>broken.paragraphs[0]!.directFontSizes()).toThrow();expect(()=>broken.paragraphs[0]!.setRunFormatting({fontSizePt:10.5})).toThrow();expect(broken.package.toBytes()).toEqual(bytes);
 }
});
test('font-size edits preserve complex-script size rather than claim to resolve it',async()=>{
 const p=await OpcPackage.open(formattingPackage());p.set(part,p.text(part).replace('<w:color w:val="008800"/>','<w:color w:val="008800"/><w:szCs w:val="30"/>'));
 const d=await Document.open(p.toBytes());d.paragraphs[0]!.setRunFormatting({fontSizePt:10.5});expect(xml(d)).toContain('<w:szCs w:val="30"/>');expect(xml(d).indexOf('w:sz ')).toBeLessThan(xml(d).indexOf('w:szCs '));
});

test('font-size handles stale after edits and late validation failures retain bytes and handles',async()=>{
 const d=await Document.open(formattingPackage()),old=d.paragraphs[0]!,span=old.find('Alpha')[0]!;old.setRunFormatting({fontSizePt:10.5});
 expect(()=>old.directFontSizes()).toThrow('stale');await expect(span.replace('Late')).rejects.toThrow('stale');
 const live=d.paragraphs[0]!,before=d.package.toBytes(),internals=d as unknown as {opcPackage:OpcPackage},serialize=internals.opcPackage.toBytes.bind(internals.opcPackage);
 internals.opcPackage.toBytes=()=>{throw Error('serialization failure');};try{expect(()=>live.setRunFormatting({fontSizePt:12})).toThrow('serialization failure');}finally{internals.opcPackage.toBytes=serialize;}
 expect(d.package.toBytes()).toEqual(before);expect(live.directFontSizes()).toEqual([10.5,10.5,10.5]);
});
test('all runs preflight font-size attributes before any early run can change',async()=>{
 const p=await OpcPackage.open(formattingPackage());p.set(part,p.text(part).replace('<w:rPr/>','<w:rPr><w:sz w:val="21" foreign="bad"/></w:rPr>'));const d=await Document.open(p.toBytes()),before=d.package.toBytes();
 expect(()=>d.paragraphs[0]!.setRunFormatting({fontSizePt:12})).toThrow();expect(d.package.toBytes()).toEqual(before);
});
test('font-size writes survive namespace aliases and UTF-16 without touching complex-script size',async()=>{
 const {utf16}=await import('../fixtures/admission.ts');const p=await OpcPackage.open(formattingPackage());p.set(part,utf16('<?xml version="1.0" encoding="UTF-16"?>'+p.text(part).replaceAll('w:', 'x:').replace('xmlns:w=','xmlns:x=')));
 const d=await Document.open(p.toBytes());d.paragraphs[0]!.setRunFormatting({fontSizePt:10.5});const bytes=d.package.get(part)!;expect([...bytes.slice(0,2)]).toEqual([255,254]);const reopened=await Document.open(d.package.toBytes());expect(reopened.paragraphs[0]!.directFontSizes()).toEqual([10.5,10.5,10.5]);expect(reopened.paragraphs[0]!.text).toBe('Alpha Beta Gamma');
});
test('font-size refuses document protection and mixed run topology without mutation',async()=>{
 for(const variant of ['protected','field']){const p=await OpcPackage.open(formattingPackage());
  if(variant==='field')p.set(part,p.text(part).replace('<w:t>Alpha</w:t>','<w:t>Alpha</w:t><w:fldChar w:fldCharType="begin"/>'));
  else{p.set('word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`);p.set('[Content_Types].xml',p.text('[Content_Types].xml').replace('</Types>','<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'));p.set('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>');}
  const d=await Document.open(p.toBytes()),before=d.package.toBytes();expect(()=>d.paragraphs[0]!.setRunFormatting({fontSizePt:10.5})).toThrow();expect(d.package.toBytes()).toEqual(before);
 }
});
test('canonical half-point scenario uses an actual saved path and reopened size',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{join}=await import('node:path'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings,cleanup}=await import('../acceptance/font-size.ts');
 const path='workflows/docx/font-size.feature',f=selectSharedScenarios(path,await Bun.file(join(fixturesRoot(),path)).text(),['@id-docx-direct-font-size-half-points']);const count=(n:number)=>({implemented:n,planned:0,total:n});
 try{const result=await executeAcceptance({root:'.',features:[f],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(6)}},bindings,'font-size-unit');expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(1);}finally{await cleanup();}
});

test('direct size API keeps its positive half-point policy separate from broader measure syntax',async()=>{
 const d=await Document.open(formattingPackage());expect(d.paragraphs[0]!.setRunFormatting({fontSizePt:0.5}).changedRuns).toBe(3);expect(d.paragraphs[0]!.directFontSizes()).toEqual([0.5,0.5,0.5]);
 const h=d.paragraphs[0]!,bytes=d.package.toBytes();expect(()=>h.setRunFormatting({fontSizePt:Number.MAX_SAFE_INTEGER})).toThrow();expect(d.package.toBytes()).toEqual(bytes);
 for(const val of ['0','12pt','21.0','9007199254740992']){const p=await OpcPackage.open(formattingPackage());p.set(part,p.text(part).replace('<w:color w:val="008800"/>',`<w:color w:val="008800"/><w:sz w:val="${val}"/>`));const opened=await Document.open(p.toBytes()),before=opened.package.toBytes();expect(()=>opened.paragraphs[0]!.directFontSizes()).toThrow();expect(()=>opened.paragraphs[0]!.setRunFormatting({fontSizePt:10.5})).toThrow();expect(opened.package.toBytes()).toEqual(before);}
});

test('font-size formatting inside a table preserves grid and existing direct properties',async()=>{
 const d=Document.create();d.addTable(1,1);d.tables[0]!.cell(0,0).text='Cell size';const opened=await Document.open(d.package.toBytes()),table=opened.tables[0]!,cell=table.cell(0,0),p=opened.paragraphs.find(p=>p.text==='Cell size')!;
 p.setRunFormatting({fontSizePt:10.5});expect(table.rows).toBe(1);expect(table.columns).toBe(1);expect(()=>cell.text).toThrow('stale');const fresh=await Document.open(opened.package.toBytes());expect(fresh.tables[0]!.cell(0,0).text).toBe('Cell size');expect(fresh.paragraphs.find(p=>p.text==='Cell size')!.directFontSizes()).toEqual([10.5]);
});

test('later malformed size blocks a combined patch and inspection without partial mutation',async()=>{
 const p=await OpcPackage.open(formattingPackage());p.set(part,p.text(part).replace('<w:rPr/>','<w:rPr><w:sz w:val="not-a-size"/></w:rPr>'));const d=await Document.open(p.toBytes()),h=d.paragraphs[0]!,before=d.package.toBytes();
 expect(()=>h.setRunFormatting({bold:true,italic:false,fontSizePt:10.5})).toThrow();expect(d.package.toBytes()).toEqual(before);expect(h.text).toBe('Alpha Beta Gamma');expect(()=>h.directFontSizes()).toThrow();
});
