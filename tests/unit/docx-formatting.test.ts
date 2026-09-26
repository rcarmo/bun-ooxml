import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings,formattingDocument,directFlags,W} from '../acceptance/run-formatting.ts';
import {Document} from '../../src/docx/index.ts';
import {OpcPackage} from '../../src/opc/package.ts';

test('direct run formatting executes15 saved outcome and refusal cases',async()=>{
 const path='workflows/docx/run-formatting.feature',f=parseFeature('references/fixtures-ooxml/'+path,(await Bun.file(join(fixturesRoot(),path)).text()).replace(/^@planned/m,'@implemented @bun')),cases=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inventory:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(cases.length),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inventory,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(15);
});
test('changed formatting stales spans and paragraphs but no-op preserves handles',async()=>{
 const d=await formattingDocument(),p=d.paragraphs[0]!,span=p.find('Alpha')[0]!;p.setRunFormatting({});expect(p.text).toBe('Alpha beta');p.setRunFormatting({bold:true});const before=await d.save();expect(()=>p.setRunFormatting({italic:true})).toThrow('stale');await expect(span.replace('Other')).rejects.toThrow('stale');expect(await d.save()).toEqual(before);
});
test('explicit false writes off properties while null removes direct overrides',async()=>{
 const d=await formattingDocument();expect(d.paragraphs[0]!.setRunFormatting({bold:false})).toEqual({changedRuns:2});expect(directFlags(new TextDecoder().decode(d.package.get('word/document.xml')!)).map(x=>x.bold)).toEqual(['0','0']);d.paragraphs[0]!.setRunFormatting({bold:null});expect(directFlags(new TextDecoder().decode(d.package.get('word/document.xml')!)).map(x=>x.bold)).toEqual([null,null]);
});

test('aliased namespaces remain bound and unrelated property bytes are retained',async()=>{
 const d=await formattingDocument(),p=await OpcPackage.open(await d.save());let xml=p.text(p.mainPart()).replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:');p.set(p.mainPart(),xml);const q=await Document.open(p.toBytes());q.paragraphs[0]!.setRunFormatting({bold:false,italic:true});
 const saved=await OpcPackage.open(await q.save()),next=saved.text(saved.mainPart());expect(next).toContain('<q:color q:val="AABBCC"/>');expect(next).toContain('<q:u q:val="single"/>');expect(directFlags(next)).toEqual([{bold:'0',italic:'1'},{bold:'0',italic:'1'}]);expect(q.paragraphs[0]!.text).toBe('Alpha beta');
});
test('reapplying explicit values preserves lexically different Boolean spellings and handles',async()=>{
 const d=Document.create();d.addParagraph('text',{bold:true,italic:true});const p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:b/>',"<w:b w:val = 'on' />").replace('<w:i/>','<w:i w:val="true"/>'));const q=await Document.open(p.toBytes()),before=await q.save(),handle=q.paragraphs[0]!;
 expect(handle.setRunFormatting({bold:true,italic:true})).toEqual({changedRuns:0});expect(await q.save()).toEqual(before);expect(handle.text).toBe('text');
});
test('absent and self-closing run properties insert in schema order',async()=>{
 for(const props of ['', '<w:rPr/>','<w:rPr data="keep"/>','<w:rPr><w:rStyle w:val="Style"/><w:rFonts w:ascii="A"/><w:color w:val="123456"/><w:u w:val="single"/></w:rPr>']){
  const d=Document.create();d.addParagraph('text');const p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:r>',`<w:r>${props}`));const q=await Document.open(p.toBytes());expect(q.paragraphs[0]!.setRunFormatting({bold:true,italic:false}).changedRuns).toBe(1);const saved=await OpcPackage.open(await q.save()),xml=saved.text(saved.mainPart());expect(directFlags(xml)).toEqual([{bold:'1',italic:'0'}]);expect(q.paragraphs[0]!.text).toBe('text');if(props.includes('data='))expect(xml).toContain('data="keep"');if(props.includes('rStyle'))expect(xml.indexOf('rStyle')).toBeLessThan(xml.indexOf('<w:b '));expect(xml.indexOf('<w:b ')).toBeLessThan(xml.indexOf('<w:i '));
 }
});
test('serialization failure retains package, earlier edits and valid paragraph handle',async()=>{
 const d=await formattingDocument();d.addParagraph('earlier edit');const before=await d.save(),handle=d.paragraphs[0]!,original=OpcPackage.prototype.toBytes;
 OpcPackage.prototype.toBytes=function(){throw new Error('injected serialization refusal');};
 try{expect(()=>handle.setRunFormatting({bold:false})).toThrow('injected serialization refusal');}finally{OpcPackage.prototype.toBytes=original;}
 expect(await d.save()).toEqual(before);expect(handle.text).toBe('Alpha beta');expect(handle.setRunFormatting({bold:false}).changedRuns).toBe(2);
});
test('UTF-16 encoding and unrelated package members survive formatting',async()=>{
 const d=await formattingDocument(),p=await OpcPackage.open(await d.save()),part=p.mainPart(),xml=p.text(part).replace('encoding="UTF-8"','encoding="UTF-16"');const bytes=new Uint8Array(2+xml.length*2),v=new DataView(bytes.buffer);v.setUint16(0,0xfeff,true);for(let i=0;i<xml.length;i++)v.setUint16(2+i*2,xml.charCodeAt(i),true);p.set(part,bytes);const q=await Document.open(p.toBytes());q.paragraphs[0]!.setRunFormatting({italic:true});const saved=await OpcPackage.open(await q.save());expect([...saved.get(part)!.slice(0,2)]).toEqual([255,254]);expect(saved.text(part)).toContain('encoding="UTF-16"');expect(directFlags(saved.text(part)).map(x=>x.italic)).toEqual(['1','1']);for(const name of p.names())if(name!==part)expect(saved.get(name)).toEqual(p.get(name));
});
test('all runs preflight before one malformed later property can change an earlier run',async()=>{
 const d=await formattingDocument(),p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:i w:val="0"/>','<w:b w:val="invalid"/><w:i w:val="0"/>'));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.paragraphs[0]!.setRunFormatting({bold:false})).toThrow();expect(await q.save()).toEqual(before);
});
test('table-cell formatting invalidates cell handles without changing grid or cell text',async()=>{
 const d=Document.create(),table=d.addTable(1,1);table.cell(0,0).text='table text';const cell=table.cell(0,0),p=d.paragraphs.find(p=>p.text==='table text')!;p.setRunFormatting({bold:true});expect(()=>cell.text).toThrow('stale');expect(table.cell(0,0).text).toBe('table text');const q=await Document.open(await d.save());expect(q.tables[0]!.rows).toBe(1);expect(q.tables[0]!.columns).toBe(1);expect(q.tables[0]!.cell(0,0).text).toBe('table text');
});
test('empty paragraphs no-op while unsupported self-closing runs refuse nonempty formatting',async()=>{
 const d=Document.create();d.addParagraph();const p=d.paragraphs[0]!,before=await d.save();expect(p.setRunFormatting({bold:true})).toEqual({changedRuns:0});expect(await d.save()).toEqual(before);
 const pkg=await OpcPackage.open(before);pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('</w:p>','<w:r/></w:p>'));const q=await Document.open(pkg.toBytes()),bytes=await q.save();expect(()=>q.paragraphs[0]!.setRunFormatting({bold:true})).toThrow();expect(await q.save()).toEqual(bytes);
});
test('unknown flags, duplicated or out-of-order properties refuse with no mutation',async()=>{
 for(const props of ['<w:b/><w:b/>','<w:color/><w:b/>','<q:b xmlns:q="urn:wrong"/>','<w:b val="0"/>','<w:rPrChange/>']){
  const d=Document.create();d.addParagraph('x');const p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:r>',`<w:r><w:rPr>${props}</w:rPr>`));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.paragraphs[0]!.setRunFormatting({bold:true})).toThrow();expect(await q.save()).toEqual(before);
 }
 const d=await formattingDocument(),before=await d.save();for(const patch of [null,[],{underline:true},{italic:'yes'}]){expect(()=>d.paragraphs[0]!.setRunFormatting(patch as any)).toThrow();expect(await d.save()).toEqual(before);}
});

test('inserting a property while deleting its following sibling yields valid ordering',async()=>{
 const d=Document.create();d.addParagraph('x',{italic:true});expect(d.paragraphs[0]!.setRunFormatting({bold:true,italic:null}).changedRuns).toBe(1);const q=await OpcPackage.open(await d.save());expect(directFlags(q.text(q.mainPart()))).toEqual([{bold:'1',italic:null}]);
});
test('default namespace elements receive qualified formatting without changing unqualified attributes',async()=>{
 const d=Document.create();d.addParagraph('x');const p=await OpcPackage.open(await d.save());let xml=p.text(p.mainPart()).replaceAll('w:','').replace('xmlns:w=', 'xmlns=');p.set(p.mainPart(),xml);const q=await Document.open(p.toBytes());q.paragraphs[0]!.setRunFormatting({bold:true,italic:false});const saved=await OpcPackage.open(await q.save());expect(directFlags(saved.text(saved.mainPart()))).toEqual([{bold:'1',italic:'0'}]);expect(q.paragraphs[0]!.text).toBe('x');
});

test('formatting does not change style graphs or Unicode text and explicitly disabled settings permit edits',async()=>{
 const {addPart,addRelationship}=await import('../../src/opc/index.ts');const d=Document.create();d.addParagraph('π 😀 café');const p=await OpcPackage.open(await d.save());
 addPart(p,'word/styles.xml',`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="BoldStyle"><w:rPr><w:b/></w:rPr></w:style></w:styles>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles','styles.xml');
 addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="false"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');
 p.set(p.mainPart(),p.text(p.mainPart()).replace(/(<w:p\b[^>]*>)/,'$1<w:pPr><w:pStyle w:val="BoldStyle"/></w:pPr>'));const q=await Document.open(p.toBytes());q.paragraphs[0]!.setRunFormatting({bold:false});const result=await OpcPackage.open(await q.save());expect(result.get('word/styles.xml')).toEqual(p.get('word/styles.xml'));expect(result.get('word/settings.xml')).toEqual(p.get('word/settings.xml'));expect(result.text(result.mainPart())).toContain('<w:pStyle w:val="BoldStyle"/>');expect(q.paragraphs[0]!.text).toBe('π 😀 café');expect(directFlags(result.text(result.mainPart()))).toEqual([{bold:'0',italic:null}]);
});
test('protected settings and mid-paragraph extension content refuse even an otherwise valid patch',async()=>{
 for(const kind of ['protected','external-settings','tracked','field']){const d=await formattingDocument(kind),before=await d.save();expect(()=>d.paragraphs[0]!.setRunFormatting({})).toThrow();expect(await d.save()).toEqual(before);}
 const d=await formattingDocument(),p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:color w:val="AABBCC"/>','<!--extension--><w:color w:val="AABBCC"/>'));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.paragraphs[0]!.setRunFormatting({bold:false})).toThrow();expect(await q.save()).toEqual(before);
});

test('fresh spans remain usable after formatting while preserving direct properties',async()=>{
 const d=await formattingDocument();d.paragraphs[0]!.setRunFormatting({bold:true});const span=d.paragraphs[0]!.find('beta')[0]!;await span.replace('gamma');const p=await OpcPackage.open(await d.save());expect(d.paragraphs[0]!.text).toBe('Alpha gamma');expect(directFlags(p.text(p.mainPart())).map(f=>f.bold)).toEqual(['1','1']);
});
test('malformed bold or italic refuses even when the other flag is requested',async()=>{
 const d=await formattingDocument(),p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:i w:val="0"/>','<w:i w:val="maybe"/>'));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.paragraphs[0]!.setRunFormatting({bold:true})).toThrow();expect(await q.save()).toEqual(before);
});
