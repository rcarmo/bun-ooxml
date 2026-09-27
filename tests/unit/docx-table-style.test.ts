import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';import {addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));
async function fixture(change=(s:string)=>s){const d=Document.create();d.addParagraph('Before');const t=d.addTable(2,2);t.cell(0,0).text='A';t.cell(1,1).text='D';const pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),change(pkg.text(pkg.mainPart())));return Document.open(pkg.toBytes());}

test('unresolved direct table style survives path reopen without registry creation or unrelated changes',async()=>{
 const d=await fixture(),t=d.tables[0]!,before=d.package.parts,held=d.paragraphs[0]!,cell=t.cell(0,0);expect(t.styleId).toBeUndefined();expect(t.setStyle('TableGrid')).toEqual({changed:1});expect(t.styleId).toBe('TableGrid');expect(d.package.has('word/styles.xml')).toBe(false);expect(()=>held.text).toThrow();expect(()=>cell.text).toThrow();expect(t.cell(0,0).text).toBe('A');expect([t.rows,t.columns]).toEqual([2,2]);
 const dir=await mkdtemp(join(tmpdir(),'table-style-'));try{const path=join(dir,'style.docx');await d.save(path);const after=await Document.open(path);expect(after.tables[0]!.styleId).toBe('TableGrid');expect(after.tables[0]!.cell(1,1).text).toBe('D');expect([...after.package.parts.keys()]).toEqual([...before.keys()]);for(const[n,b]of before)if(n!=='word/document.xml')expect(after.package.get(n)).toEqual(b);}finally{await rm(dir,{recursive:true,force:true});}
 expect(elements(parseXml(xml(d)),'tblPr',W)[0]!.children.map(n=>n.localName)).toEqual(['tblStyle','tblW']);
});

test('semantic no-ops retain lexical style ID spelling and handles; null removes only the direct reference',async()=>{
 const raw='<w:tblStyle w:val = \'A&#x26;B\'/>',d=await fixture(x=>x.replace('<w:tblPr>','<w:tblPr>'+raw)),t=d.tables[0]!,cell=t.cell(0,0),before=d.package.toBytes();expect(t.styleId).toBe('A&B');expect(t.setStyle('A&B')).toEqual({changed:0});expect(d.package.toBytes()).toEqual(before);expect(cell.text).toBe('A');const source=xml(d);expect(t.setStyle(null)).toEqual({changed:1});expect(t.styleId).toBeUndefined();expect(xml(d)).toBe(source.replace(raw,''));const noStyle=d.package.toBytes();expect(t.setStyle(null)).toEqual({changed:0});expect(d.package.toBytes()).toEqual(noStyle);
});

test('style writes preserve table properties, all existing row fragments and an unrelated styles registry',async()=>{
 const d=await fixture();d.addParagraphStyle('Unrelated',{name:'Other',bold:true});const before=d.package.parts,source=xml(d),rows=elements(parseXml(source),'tr',W).map(r=>source.slice(r.start,r.end));d.tables[0]!.setStyle('NotDefined');expect(d.tables[0]!.styleId).toBe('NotDefined');for(const raw of rows)expect(xml(d)).toContain(raw);expect(xml(d).replace(/<w:tblStyle[^>]*\/>/,'')).toBe(source);for(const[n,b]of before)if(n!=='word/document.xml')expect(d.package.get(n)).toEqual(b);
});

test('missing or self-closing table properties accept a first-position direct reference',async()=>{
 for(const replacement of ['', '<w:tblPr/>']){const d=await fixture(x=>x.replace(/<w:tblPr>[^]*?<\/w:tblPr>/,replacement)),t=d.tables[0]!;expect(t.setStyle('TableGrid')).toEqual({changed:1});expect(t.styleId).toBe('TableGrid');expect((await Document.open(d.package.toBytes())).tables[0]!.styleId).toBe('TableGrid');expect(t.cell(0,0).text).toBe('A');}
});

test('invalid style arguments refuse with unchanged bytes and current handles',async()=>{
 const d=await fixture(),t=d.tables[0]!,before=d.package.toBytes();for(const value of [undefined,'',' ',' Leading','Trailing ','a\tb','a\u0000b','a\u0085b','\ud800','x'.repeat(256),1,{},new String('TableGrid')]){expect(()=>t.setStyle(value as string)).toThrow();expect(d.package.toBytes()).toEqual(before);expect(t.styleId).toBeUndefined();}expect(t.setStyle('雪 & <style>')).toEqual({changed:1});expect(t.styleId).toBe('雪 & <style>');expect((await Document.open(d.package.toBytes())).tables[0]!.styleId).toBe('雪 & <style>');
});

test('duplicate misordered revised decorated or wrong-namespace style metadata refuses even removal or no-op',async()=>{
 for(const pr of ['<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblStyle w:val="Other"/></w:tblPr>','<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblStyle w:val="TableGrid"/></w:tblPr>','<w:tblPr><w:tblStyle val="TableGrid"/></w:tblPr>','<w:tblPr><w:tblStyle w:val="TableGrid" w:extra="1"/></w:tblPr>','<w:tblPr><w:tblStyle w:val=""/></w:tblPr>','<w:tblPr><w:tblStyle w:val="TableGrid"><w:bad/></w:tblStyle></w:tblPr>','<w:tblPr><!--keep--><w:tblStyle w:val="TableGrid"/></w:tblPr>','<w:tblPr><w:tblPrChange/></w:tblPr>','<w:tblPr><w:unknown/></w:tblPr>','<w:tblPr/><w:tblPr/>']){const d=await fixture(x=>x.replace(/<w:tblPr>[^]*?<\/w:tblPr>/,pr)),before=d.package.toBytes(),t=d.tables[0]!;expect(()=>t.styleId).toThrow();expect(()=>t.setStyle('TableGrid')).toThrow();expect(()=>t.setStyle(null)).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('merged revised nested or unsupported row topology refuses style writes with custody intact',async()=>{
 for(const change of [(x:string)=>x.replace('<w:tcPr>','<w:tcPr><w:vMerge/>'),(x:string)=>x.replace('<w:tr>','<w:tr><w:trPr><w:ins/></w:trPr>'),(x:string)=>x.replace('<w:t>A</w:t>','<w:fldChar w:fldCharType="begin"/>'),(x:string)=>x.replace('</w:tc>','<w:tbl><w:tblGrid><w:gridCol w:w="10"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc>'),(x:string)=>x.replace(/<w:tblGrid>[^]*?<\/w:tblGrid>/,''),(x:string)=>x.replace('</w:tbl>','<w:tblPr/></w:tbl>')]){const d=await fixture(change),before=d.package.toBytes();expect(()=>d.tables[0]!.setStyle('TableGrid')).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('protection, external main XML and structural staleness refuse while protected direct reads stay available',async()=>{
 const d=await fixture(x=>x.replace('<w:tblPr>','<w:tblPr><w:tblStyle w:val="TableGrid"/>')),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(pkg.toBytes()),before=locked.package.toBytes();expect(locked.tables[0]!.styleId).toBe('TableGrid');expect(()=>locked.tables[0]!.setStyle('TableGrid')).toThrow();expect(locked.package.toBytes()).toEqual(before);
 const t=d.tables[0]!;d.package.setPart('word/document.xml',new TextEncoder().encode(xml(d).replace('Before','External')));const changed=d.package.toBytes();expect(()=>t.styleId).toThrow();expect(()=>t.setStyle(null)).toThrow();expect(d.package.toBytes()).toEqual(changed);
 const fresh=await fixture(),old=fresh.tables[0]!;fresh.addParagraph('new');expect(()=>old.styleId).toThrow();expect(()=>old.setStyle('TableGrid')).toThrow();
});

test('transactional set and serialization faults retain table paragraph and cell handles',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=await fixture(),t=d.tables[0]!,p=d.paragraphs[0]!,c=t.cell(0,0),before=d.package.toBytes(),pkg=(d as any).opcPackage as OpcPackage,original=pkg[stage].bind(pkg);if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialize');};try{expect(()=>t.setStyle('TableGrid')).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}expect(d.package.toBytes()).toEqual(before);expect(t.styleId).toBeUndefined();expect(p.text).toBe('Before');expect(c.text).toBe('A');}
});

test('alias and UTF16/BOM edits preserve source namespace and encoding after byte reopen',async()=>{
 for(const encoding of ['utf8','utf16']){const d=await fixture(x=>x.replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('<q:tblPr>','<q:tblPr xmlns:w="urn:foreign">')),pkg=await OpcPackage.open(d.package.toBytes());let raw=pkg.get(pkg.mainPart())!;if(encoding==='utf8')raw=new Uint8Array([239,187,191,...raw]);else{const text=pkg.text(pkg.mainPart()).replace('UTF-8','UTF-16');raw=new Uint8Array(2+text.length*2);raw.set([254,255]);const view=new DataView(raw.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+2*i,text.charCodeAt(i),false);}pkg.set(pkg.mainPart(),raw);const doc=await Document.open(pkg.toBytes());doc.tables[0]!.setStyle('TableGrid');const after=await Document.open(await doc.save());expect(after.tables[0]!.styleId).toBe('TableGrid');expect(after.tables[0]!.cell(0,0).text).toBe('A');expect([...after.package.get('word/document.xml')!.slice(0,encoding==='utf8'?3:2)]).toEqual(encoding==='utf8'?[239,187,191]:[254,255]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(after.package.get(n)).toEqual(pkg.get(n));}
});

test('canonical table-style readback rejects false initial or changed style values',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/table-style.ts');const path='workflows/docx/tables.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(3)}});
 const good=await executeAcceptance(inv(source),bindings,'table-style-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);const wrong=await executeAcceptance(inv(source.replace('the first style is empty and the second style is TableGrid','the first style is empty and the second style is Wrong')),bindings,'table-style-expected');expect(wrong.counts.cases.failed).toBe(1);expect(wrong.counts.steps.failed).toBe(1);expect(wrong.counts.steps.undefined).toBe(0);
 for(const field of ['styleBefore','styleAfter']){const corrupt=bindings.map(b=>b.pattern.test('its style is read, then set to TableGrid and read again')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);(c.state as Record<string,unknown>)[field]='Corrupt';}}:b);const bad=await executeAcceptance(inv(source),corrupt,'table-style-corrupt');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
});

test('complete output bounds refuse adding a style reference without invalidating handles',async()=>{
 const d=await fixture(),p=await OpcPackage.open(d.package.toBytes()),main=p.mainPart(),source=p.text(main),padding=' '.repeat(8*1024*1024-source.length-2);p.set(main,source.replace('</w:body>',padding+'</w:body>'));const loaded=await Document.open(p.toBytes()),t=loaded.tables[0]!,c=t.cell(0,0),before=loaded.package.toBytes();expect(()=>t.setStyle('TableGrid')).toThrow(expect.objectContaining({code:'XML_EDIT_UNSAFE'}));expect(loaded.package.toBytes()).toEqual(before);expect(t.styleId).toBeUndefined();expect(c.text).toBe('A');
});
