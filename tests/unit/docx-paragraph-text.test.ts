import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';
import {addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,applyEdits} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const text=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));
async function fixture(content:string){const d=Document.create();d.addParagraph('Alpha');d.addParagraph('Sibling');const pkg=await OpcPackage.open(d.package.toBytes()),xml=pkg.text(pkg.mainPart()),p=elements(parseXml(xml),'p',W)[0]!;pkg.set(pkg.mainPart(),applyEdits(xml,[{start:p.openEnd,end:p.closeStart,value:content}]));return Document.open(pkg.toBytes());}

test('whole paragraph text replaces empty and existing values with fresh handles and saved readback',async()=>{
 const d=Document.create();let p=d.addParagraph('');
 for(const value of ['', 'Hello World','  spaces  ','日本語テキスト','a < b > c & d','雪 😀']){const held=p,old=p.text,before=d.package.toBytes();p=p.setText(value);expect(p.text).toBe(value);expect(p).toBe(d.paragraphs[0]!);if(value===old){expect(p).toBe(held);expect(d.package.toBytes()).toEqual(before);}else expect(()=>held.text).toThrow();const reopened=await Document.open(d.package.toBytes());expect(reopened.paragraphs[0]!.text).toBe(value);}
 p=p.setText('');expect(p.text).toBe('');expect(p.directRunFlags()).toHaveLength(1);
 const dir=await mkdtemp(join(tmpdir(),'paragraph-text-'));try{const path=join(dir,'text.docx');p=p.setText('  saved <雪>  ');await d.save(path);expect((await Document.open(path)).paragraphs[0]!.text).toBe('  saved <雪>  ');}finally{await rm(dir,{recursive:true,force:true});}
});

test('replacement keeps all paragraph/run properties and sibling/member fragments but concentrates text in first leaf',async()=>{
 const ppr='<w:pPr><w:keepNext/></w:pPr>',rpr='<w:rPr><w:b/><w:color w:val = \'123456\'/></w:rPr>';
 const d=await fixture(`${ppr}<w:r w:rsidR="00112233">${rpr}<w:t>A</w:t><w:t>B</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>C</w:t></w:r>`),before=d.package.parts,p=d.paragraphs[0]!;
 const sibling=elements(parseXml(text(d)),'p',W)[1]!,raw=text(d).slice(sibling.start,sibling.end);const next=p.setText(' New & <雪> '),xml=text(d);expect(next.text).toBe(' New & <雪> ');expect(xml).toContain(ppr);expect(xml).toContain(rpr);expect(xml).toContain('<w:r w:rsidR="00112233">');expect(xml).toContain('<w:rPr><w:i/></w:rPr>');expect(xml).toContain(raw);
 expect(elements(parseXml(xml),'t',W).slice(0,3).map(n=>n.text)).toEqual([' New & <雪> ','','']);expect(next.directRunFlags().map(f=>[f.bold,f.italic])).toEqual([[true,null],[null,true]]);expect(xml).toContain('xml:space="preserve"');const reopened=await Document.open(d.package.toBytes());expect(reopened.paragraphs[0]!.text).toBe(' New & <雪> ');for(const[n,b]of before)if(n!=='word/document.xml')expect(reopened.package.get(n)).toEqual(b);
});

test('same text is a validated lexical no-op retaining entities, archive and held span',async()=>{
 const d=await fixture('<w:r><w:t>A&#x26;B</w:t></w:r>'),p=d.paragraphs[0]!,span=p.find('A')[0]!,before=d.package.toBytes();expect(p.setText('A&B')).toBe(p);expect(d.package.toBytes()).toEqual(before);await span.replace('X');expect(d.paragraphs[0]!.text).toBe('X&B');
});

test('self-closing paragraphs, runs and text leaves insert safely under alias and foreign w bindings',async()=>{
 for(const content of ['<q:p xmlns:q="'+W+'" xmlns:w="urn:foreign"/>','<q:p xmlns:q="'+W+'" xmlns:w="urn:foreign"><q:pPr/><q:r><q:rPr><q:b/></q:rPr></q:r></q:p>','<q:p xmlns:q="'+W+'"><q:r/><q:r><q:t/></q:r></q:p>','<q:p xmlns:q="'+W+'" xmlns:w="urn:foreign"><q:r/></q:p>']){const d=Document.create(),pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('<w:body>','<w:body>'+content));const loaded=await Document.open(pkg.toBytes()),p=loaded.paragraphs[0]!.setText('one');expect(p.text).toBe('one');expect((await Document.open(loaded.package.toBytes())).paragraphs[0]!.text).toBe('one');expect(elements(parseXml(text(loaded)),'t','urn:foreign')).toHaveLength(0);}
});

test('invalid text types, XML characters, controls and input limits refuse without changing handles or bytes',()=>{
 const d=Document.create(),p=d.addParagraph('Old'),before=d.package.toBytes();for(const value of [null,undefined,1,{},new String('x'),'x\t','x\n','x\r','\u0000','\ud800','x'.repeat(1024*1024+1)]){expect(()=>p.setText(value as string)).toThrow();expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Old');}
});

test('fields, revisions, lexical text barriers and unknown text attributes refuse even on same text',async()=>{
 for(const content of ['<w:hyperlink><w:r><w:t>Alpha</w:t></w:r></w:hyperlink>','<w:r><w:fldChar w:fldCharType="begin"/></w:r>','<w:r><w:t>Al<!--keep-->pha</w:t></w:r>','<w:r><w:t><![CDATA[Alpha]]></w:t></w:r>','<w:r><w:t><?x y?>Alpha</w:t></w:r>','<w:r><w:t w:unknown="keep">Alpha</w:t></w:r>','<w:r><w:t xml:space="bad">Alpha</w:t></w:r>','<w:pPr><w:pPrChange/></w:pPr><w:r><w:t>Alpha</w:t></w:r>','<w:r><w:rPr><w:rPrChange/></w:rPr><w:t>Alpha</w:t></w:r>','<w:r><w:t>Alpha</w:t></w:r><!--keep-->']){const d=await fixture(content),before=d.package.toBytes();for(const value of ['Alpha','New',''])expect(()=>d.paragraphs[0]!.setText(value)).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('protection, stale handles and externally changed source refuse before write or no-op',async()=>{
 const d=Document.create();d.addParagraph('Old');const pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(pkg.toBytes()),lb=locked.package.toBytes();expect(()=>locked.paragraphs[0]!.setText('Old')).toThrow();expect(locked.package.toBytes()).toEqual(lb);
 const held=d.paragraphs[0]!;held.setText('New');expect(()=>held.setText('New')).toThrow(expect.objectContaining({code:'docx-stale-paragraph'}));const fresh=d.paragraphs[0]!;d.package.setPart('word/document.xml',new TextEncoder().encode(text(d).replace('New','External')));const external=d.package.toBytes();expect(()=>fresh.setText('New')).toThrow();expect(d.package.toBytes()).toEqual(external);
});

test('write and serialization faults roll back bytes while retaining paragraph/span/cell handles',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=Document.create();d.addTable(1,1);d.tables[0]!.cell(0,0).text='Old';const p=d.paragraphs[0]!,cell=d.tables[0]!.cell(0,0),span=p.find('Old')[0]!,before=d.package.toBytes(),pkg=(d as unknown as {opcPackage:OpcPackage}).opcPackage,original=pkg[stage].bind(pkg);
 if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialize');};try{expect(()=>p.setText('New')).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Old');expect(cell.text).toBe('Old');await span.replace('Works');expect(d.paragraphs[0]!.text).toBe('Works');}
});

test('table paragraph changes retain dimensions and sibling text with UTF-16 byte reopen',async()=>{
 const d=Document.create(),table=d.addTable(1,2);table.cell(0,0).text='Old';table.cell(0,1).text='Sibling';const pkg=await OpcPackage.open(d.package.toBytes()),xml=pkg.text(pkg.mainPart()).replace('UTF-8','UTF-16'),bytes=new Uint8Array(2+xml.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<xml.length;i++)view.setUint16(2+2*i,xml.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);const loaded=await Document.open(pkg.toBytes()),t=loaded.tables[0]!,c=t.cell(0,0),span=loaded.paragraphs[0]!.find('Old')[0]!;loaded.paragraphs[0]!.setText('New 雪');expect([t.rows,t.columns]).toEqual([1,2]);expect(t.cell(0,1).text).toBe('Sibling');expect(()=>c.text).toThrow();await expect(span.replace('bad')).rejects.toThrow(expect.objectContaining({code:'docx-stale-span'}));
 const reopened=await Document.open(loaded.package.toBytes());expect(reopened.tables[0]!.cell(0,0).text).toBe('New 雪');expect(reopened.tables[0]!.cell(0,1).text).toBe('Sibling');expect([...reopened.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);
});

test('xml:space edits retain declaration values, quotes and encoded default while preserving first-leaf boundary spaces',async()=>{
 for(const attrs of ['xml:space=\'de&#x66;ault\'','xmlns:x=\'urn: xml:space="default"\' xml:space="default"','xmlns:x="urn: xml:space=\'default\'" xml:space=\'default\'','xmlns:x="urn:test"']){
  const d=await fixture(`<w:r><w:t ${attrs}> a </w:t><w:t>tail</w:t></w:r>`),next=d.paragraphs[0]!.setText(' a '),xml=text(d),leaves=elements(parseXml(xml),'t',W);expect(next.text).toBe(' a ');expect(leaves[0]!.attributes['xml:space']).toBe('preserve');expect(leaves[1]!.text).toBe('');const declaration=attrs.match(/xmlns:x=(?:"[^"]*"|'[^']*')/)?.[0];if(declaration)expect(xml).toContain(declaration);expect((await Document.open(d.package.toBytes())).paragraphs[0]!.text).toBe(' a ');
 }
});

test('setter output bounds retain archive and live handles',async()=>{
 const d=Document.create();d.addParagraph('');const pkg=await OpcPackage.open(d.package.toBytes()),main=pkg.mainPart(),xml=pkg.text(main),padding='x'.repeat(8*1024*1024-xml.length-9);pkg.set(main,xml.replace('</w:body>','<!--'+padding+'--></w:body>'));const loaded=await Document.open(pkg.toBytes()),p=loaded.paragraphs[0]!,before=loaded.package.toBytes();expect(()=>p.setText('x')).toThrow(expect.objectContaining({code:'XML_EDIT_UNSAFE'}));expect(loaded.package.toBytes()).toEqual(before);expect(p.text).toBe('');
});

test('five canonical text getter cases execute and corrupted getter/expected values fail assertions',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/paragraph-text.ts');const path='workflows/docx/document-model.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(5),steps:count(15)}});
 const good=await executeAcceptance(inv(source),bindings,'paragraph-text-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(5);
 const bad=await executeAcceptance(inv(source.replace('the paragraph text getter equals JSON <text_json>','the paragraph text getter equals JSON "wrong"')),bindings,'paragraph-text-expected');expect(bad.counts.cases.failed).toBe(5);expect(bad.counts.steps.failed).toBe(5);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);
 const corrupt=bindings.map(b=>b.pattern.test('its text is set to JSON "x"')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);(c.state as {document:Document}).document.paragraphs[0]!.setText('CORRUPT');}}:b),wrong=await executeAcceptance(inv(source),corrupt,'paragraph-text-corrupt');expect(wrong.counts.cases.failed).toBe(5);expect(wrong.counts.steps.failed).toBe(5);expect(wrong.counts.steps.undefined).toBe(0);expect(wrong.counts.steps.ambiguous).toBe(0);
});

test('first existing text leaf owns replacement even after an empty run and no-leaf insertion retains first-run formatting',async()=>{
 const later=await fixture('<w:r><w:rPr><w:b/></w:rPr></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Old</w:t></w:r>');const p=later.paragraphs[0]!.setText('New'),parsed=parseXml(text(later)),runs=elements(parsed,'r',W);expect(p.text).toBe('New');expect(runs[0]!.children.some(n=>n.localName==='t')).toBe(false);expect(runs[1]!.children.find(n=>n.localName==='t')!.text).toBe('New');expect(p.directRunFlags().map(f=>[f.bold,f.italic])).toEqual([[true,null],[null,true]]);
 const absent=await fixture('<w:r><w:rPr><w:b/></w:rPr></w:r><w:r><w:rPr><w:i/></w:rPr></w:r>');absent.paragraphs[0]!.setText('First');const nodes=elements(parseXml(text(absent)),'r',W);expect(nodes[0]!.children.find(n=>n.localName==='t')!.text).toBe('First');expect(nodes[1]!.children.some(n=>n.localName==='t')).toBe(false);
});
