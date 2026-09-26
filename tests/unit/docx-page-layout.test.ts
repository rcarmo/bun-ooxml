import {test,expect} from 'bun:test';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/page-layout.ts';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
test('final section geometry executes22 saved outcomes and atomic refusals',async()=>{
 const path='workflows/docx/page-layout.feature',f=parseFeature('references/fixtures-ooxml/'+path,(await Bun.file(join(fixturesRoot(),path)).text()).replace(/^@planned/m,'@implemented @bun')),rows=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(rows.length),steps:count(rows.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inv,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(22);
});

import {Document,W_NS as W} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {layoutDocument,desired,originalLayout} from '../acceptance/page-layout.ts';
import {parseXml,elements} from '../../src/xml/index.ts';

test('page geometry changes stale paragraphs and cells while table handles remain valid',async()=>{
 const d=await layoutDocument(),table=d.addTable(1,1);table.cell(0,0).text='cell';const p=d.paragraphs[0]!,span=p.find('Layout')[0]!,cell=table.cell(0,0),before=await d.save();expect(d.getPageLayout()).toEqual(originalLayout);expect(await d.save()).toEqual(before);d.setPageLayout(desired('landscape'));expect(()=>p.text).toThrow('stale');expect(()=>cell.text).toThrow('stale');await expect(span.replace('wrong')).rejects.toThrow('stale');expect(table.cell(0,0).text).toBe('cell');const q=await Document.open(await d.save());expect(q.getPageLayout()).toEqual(desired('landscape'));expect(q.tables[0]!.cell(0,0).text).toBe('cell');
});
test('numeric and default-orientation no-op retains exact spelling and existing handles',async()=>{
 const d=await layoutDocument(),p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('w:w="12240"',"w:w = '012240'"));const q=await Document.open(p.toBytes()),before=await q.save(),handle=q.paragraphs[0]!;const copy=q.getPageLayout();copy.width=3;expect(q.getPageLayout().width).toBe(12240);expect(q.setPageLayout(originalLayout)).toEqual({changed:0});expect(await q.save()).toEqual(before);expect(q.paragraphs[0]).toBe(handle);
});
test('layout write preserves UTF16 LE and BE encoding, BOM and body text',async()=>{
 for(const little of [true,false]){const d=await layoutDocument(),p=await OpcPackage.open(await d.save()),xml=p.text(p.mainPart()).replace('encoding="UTF-8"','encoding="UTF-16"'),bytes=new Uint8Array(2+xml.length*2),view=new DataView(bytes.buffer);view.setUint16(0,0xfeff,little);for(let i=0;i<xml.length;i++)view.setUint16(2+i*2,xml.charCodeAt(i),little);p.set(p.mainPart(),bytes);const q=await Document.open(p.toBytes());q.setPageLayout(desired('landscape'));const saved=await OpcPackage.open(await q.save());expect([...saved.get(saved.mainPart())!.slice(0,2)]).toEqual(little?[255,254]:[254,255]);expect(saved.text(saved.mainPart())).toContain('encoding="UTF-16"');expect(q.paragraphs[0]!.text).toBe('Layout text π');}
});
test('serialization refusal leaves prior changes, page values and paragraph identity intact',async()=>{
 const d=await layoutDocument();d.addParagraph('earlier');const before=await d.save(),handle=d.paragraphs[0]!,serialize=OpcPackage.prototype.toBytes;OpcPackage.prototype.toBytes=function(){throw Error('injected layout serialization');};try{expect(()=>d.setPageLayout(desired('landscape'))).toThrow('injected layout serialization');}finally{OpcPackage.prototype.toBytes=serialize;}expect(await d.save()).toEqual(before);expect(d.getPageLayout()).toEqual(originalLayout);expect(d.paragraphs[0]).toBe(handle);d.setPageLayout(desired('landscape'));expect(d.paragraphs.at(-1)!.text).toBe('earlier');
});
test('layout options reject getters, incomplete fields, unknown keys and oversized geometry',async()=>{
 const d=await layoutDocument(),before=await d.save();let calls=0;for(const opts of [{...originalLayout,get width(){calls++;return 1;}},{width:10},{...originalLayout,unknown:true},{...originalLayout,width:31681},{...originalLayout,width:1.5},{...originalLayout,width:NaN},{...originalLayout,orientation:'landscape'},{...originalLayout,header:20000},{...originalLayout,[Symbol('extra')]:1}]){expect(()=>d.setPageLayout(opts as any)).toThrow();expect(await d.save()).toEqual(before);}expect(calls).toBe(0);
});
test('orientation writes avoid shadowed namespace prefixes and retain paper codes and attributes',async()=>{
 const d=await layoutDocument('aliased'),p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('<q:pgSz ','<q:pgSz xmlns:layout="urn:other" q:code="9" '));const q=await Document.open(p.toBytes());q.setPageLayout(desired('landscape'));const xml=(await OpcPackage.open(await q.save())).text('word/document.xml');expect(xml).toContain('q:code="9"');expect(xml).toContain('xmlns:layout="urn:other"');expect(xml).toContain('layoutx:orient="landscape"');expect(q.getPageLayout()).toEqual(desired('landscape'));
});
test('mirror/bookfold/gutter settings and section bidi flags refuse conservatively',async()=>{
 for(const name of ['mirrorMargins','gutterAtTop','bookFoldPrinting','bookFoldRevPrinting','bidi','rtlGutter']){const d=await layoutDocument(),p=await OpcPackage.open(await d.save());if(name==='bidi'||name==='rtlGutter')p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:docGrid',`<w:${name}/><w:docGrid`));else{addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:${name}/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');}const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.setPageLayout(desired('landscape'))).toThrow();expect(await q.save()).toEqual(before);}
});
test('protection applies to same-geometry requests and disabled settings allow mutation',async()=>{
 const d=await layoutDocument('protected'),before=await d.save();expect(()=>d.setPageLayout(originalLayout)).toThrow();expect(await d.save()).toEqual(before);const p=await OpcPackage.open(before);p.set('word/settings.xml',p.text('word/settings.xml').replace('enforcement="1"','enforcement="0"'));const q=await Document.open(p.toBytes());q.setPageLayout(desired('landscape'));expect(q.getPageLayout().orientation).toBe('landscape');
});
test('invalid metadata geometry attributes and page-size children refuse reads and mutations',async()=>{
 for(const [from,to]of [['w:top="1440"','w:top="-1"'],['w:header="720"',''],['w:h="15840"','w:h="12px"'],['w:w="12240"','w:w="12240" w:orient="sideways"'],['<w:pgSz w:w="12240" w:h="15840"/>','<w:pgSz w:w="12240" w:h="15840"><w:bad/></w:pgSz>']] as const){const d=await layoutDocument(),p=await OpcPackage.open(await d.save()),xml=p.text(p.mainPart());expect(xml.includes(from)).toBe(true);p.set(p.mainPart(),xml.replace(from,to));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.getPageLayout()).toThrow();expect(()=>q.setPageLayout(desired('landscape'))).toThrow();expect(await q.save()).toEqual(before);}
});
test('disk save retains final geometry and leaves earlier section XML byte-identical',async()=>{
 const {mkdtemp,rm}=await import('node:fs/promises'),{tmpdir}=await import('node:os'),{join}=await import('node:path'),root=await mkdtemp(join(tmpdir(),'bun-page-layout-'));try{const d=await layoutDocument('earlier-section'),before=new TextDecoder().decode(d.package.get('word/document.xml')!),section=elements(parseXml(before),'sectPr',W)[0]!,raw=before.slice(section.start,section.end);d.setPageLayout(desired('landscape'));const path=join(root,'landscape.docx');await d.save(path);const q=await Document.open(path);expect(q.getPageLayout()).toEqual(desired('landscape'));expect(new TextDecoder().decode(q.package.get('word/document.xml')!)).toContain(raw);}finally{await rm(root,{recursive:true,force:true});}
});

test('margin child elements and text refuse reads and writes without mutation',async()=>{
 for(const content of ['<w:bad/>','unexpected text']){const d=await layoutDocument(),p=await OpcPackage.open(await d.save()),xml=p.text(p.mainPart());p.set(p.mainPart(),xml.replace(/(<w:pgMar[^>]*?)\/>/,`$1>${content}</w:pgMar>`));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.getPageLayout()).toThrow();expect(()=>q.setPageLayout(desired('landscape'))).toThrow();expect(await q.save()).toEqual(before);}
});
test('attribute-only layout edits preserve exact unrelated spelling inside geometry nodes',async()=>{
 const d=await layoutDocument(),p=await OpcPackage.open(await d.save());const xml=p.text(p.mainPart()).replace('<w:pgSz w:w="12240" w:h="15840"/>',`<w:pgSz xmlns:z='urn:custom' w:code = '009' w:w = '12240' w:h="15840" />`).replace('w:header="720"',"w:header = '00720'");p.set(p.mainPart(),xml);const q=await Document.open(p.toBytes());q.setPageLayout({...originalLayout,width:12000,top:1200});const next=(await OpcPackage.open(await q.save())).text(p.mainPart());expect(next).toBe(xml.replace("w:w = '12240'",'w:w = "12000"').replace('w:top="1440"','w:top="1200"'));
});
