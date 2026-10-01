import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';import {addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,attribute,type XmlElement} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const pkg=(d:Document)=>(d as any).opcPackage as OpcPackage;
const xml=(d:Document)=>pkg(d).text(pkg(d).mainPart());
async function fixture(){const d=Document.create();d.addParagraph('Before');d.addTable(3,3);d.tables[0]!.cell(0,1).text='Keep & <text>';d.tables[0]!.cell(0,1).setProperties({shading:'ABCDEF',verticalAlign:'center'});d.tables[0]!.cell(2,2).text='Tail';d.addParagraph('After');const p=await OpcPackage.open(d.package.toBytes());addPart(p,'customXml/opaque.bin',new Uint8Array([0,255,42]),'application/octet-stream');return Document.open(p.toBytes());}
function cells(source:string){return elements(parseXml(source),'tc',W);}
function replaceCell(p:OpcPackage,index:number,change:(s:string)=>string){const source=p.text(p.mainPart()),c=cells(source)[index]!;p.set(p.mainPart(),source.slice(0,c.start)+change(source.slice(c.start,c.end))+source.slice(c.end));}
function withoutFlags(source:string){let result=source;for(const n of elements(parseXml(source),'vMerge',W).sort((a,b)=>b.start-a.start))result=result.slice(0,n.start)+result.slice(n.end);return result;}

test('vertical merge saves restart and explicit continue markers without removing cells text widths or formatting',async()=>{
 const d=await fixture(),before=d.package.parts,source=xml(d),paragraphs=d.paragraphs.map(p=>p.text),held=d.tables[0]!,cell=held.cell(0,1),paragraph=d.paragraphs[0]!,current=held.mergeColumnCells(1,0,2);
 expect(current.rows).toBe(3);expect(current.columns).toBe(3);expect(()=>held.rows).toThrow(expect.objectContaining({code:'docx-stale-table'}));expect(()=>cell.text).toThrow();expect(()=>paragraph.text).toThrow();for(const row of [0,1,2])expect(()=>current.cell(row,1)).toThrow(expect.objectContaining({code:'docx-table-merged-cell'}));expect(current.cell(2,2).text).toBe('Tail');
 const next=xml(d),tree=parseXml(next);expect(elements(tree,'tc',W)).toHaveLength(9);expect(elements(tree,'vMerge',W).map(n=>attribute(n,'val',W))).toEqual(['restart','continue','continue']);expect(withoutFlags(next)).toBe(source);expect(d.paragraphs.map(p=>p.text)).toEqual(paragraphs);
 const dir=await mkdtemp(join(tmpdir(),'vertical-merge-'));try{const path=join(dir,'merged.docx');await d.save(path);const read=await Document.open(path);expect(read.paragraphs.map(p=>p.text)).toEqual(paragraphs);expect(read.tables[0]!.cell(2,2).text).toBe('Tail');expect(elements(parseXml(xml(read)),'vMerge',W).map(n=>attribute(n,'val',W))).toEqual(['restart','continue','continue']);for(const[n,b]of before)if(n!=='word/document.xml')expect(read.package.get(n)).toEqual(b);expect([...read.package.parts.keys()]).toEqual([...before.keys()]);}finally{await rm(dir,{recursive:true,force:true});}
});

test('middle and edge ranges preserve physical geometry and all paragraphs',async()=>{
 for(const[col,first,last]of [[0,0,1],[0,1,2],[2,0,1]]){const d=await fixture(),source=xml(d),before=d.paragraphs.map(p=>p.text),t=d.tables[0]!.mergeColumnCells(col!,first!,last!);expect(t.rows).toBe(3);expect(t.columns).toBe(3);expect(withoutFlags(xml(d))).toBe(source);expect(d.paragraphs.map(p=>p.text)).toEqual(before);expect(elements(parseXml(xml(d)),'vMerge',W)).toHaveLength(2);}
});

test('continuation content and decorations refuse rather than being hidden or discarded',async()=>{
 for(const kind of ['text','space','format','paragraph-format','empty-run','lexical','attribute']){const d=await fixture(),p=await OpcPackage.open(d.package.toBytes());replaceCell(p,4,raw=>kind==='text'||kind==='space'?raw.replace(/<w:p\b[^>]*\/>/,`<w:p><w:r><w:t xml:space="preserve">${kind==='text'?'Do not hide me':' '}</w:t></w:r></w:p>`):kind==='format'?raw.replace('</w:tcPr>','<w:shd w:val="clear" w:fill="ABCDEF"/></w:tcPr>'):kind==='paragraph-format'?raw.replace(/<w:p\b[^>]*\/>/,'<w:p><w:pPr><w:keepNext/></w:pPr></w:p>'):kind==='empty-run'?raw.replace(/<w:p\b[^>]*\/>/,'<w:p><w:r><w:t/></w:r></w:p>'):kind==='attribute'?raw.replace('<w:tc>','<w:tc keep="yes">'):raw.replace('<w:tcPr>','<!--keep--><w:tcPr>'));const loaded=await Document.open(p.toBytes()),before=loaded.package.toBytes(),t=loaded.tables[0]!;expect(()=>t.mergeColumnCells(1,0,2)).toThrow();expect(loaded.package.toBytes()).toEqual(before);expect(t.rows).toBe(3);}
});

test('invalid vertical coordinates refuse atomically without coercion',async()=>{
 const d=await fixture(),t=d.tables[0]!,before=d.package.toBytes();let invoked=false;for(const args of [[1,0,0],[1,2,1],[-1,0,2],[3,0,2],[1,-1,2],[1,0,3],[1,0.5,2],[NaN,0,2],['1',0,2],[1,0,Infinity],[1,0,Number.MAX_SAFE_INTEGER+1],[{valueOf(){invoked=true;return 1;}},0,2]]){expect(()=>t.mergeColumnCells(...args as [number,number,number])).toThrow(RangeError);expect(d.package.toBytes()).toEqual(before);}expect(invoked).toBe(false);
});

test('whole-table existing merges nested cells invalid widths and revised metadata refuse even outside selected column',async()=>{
 for(const kind of ['horizontal','vertical','nested','width','revision','missing-grid','late-span']){const d=await fixture(),p=await OpcPackage.open(d.package.toBytes());if(kind==='missing-grid')p.set(p.mainPart(),p.text(p.mainPart()).replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/,''));else replaceCell(p,8,raw=>kind==='horizontal'?raw.replace('<w:tcPr>','<w:tcPr><w:gridSpan w:val="2"/>'):kind==='vertical'?raw.replace('<w:tcPr>','<w:tcPr><w:vMerge w:val="restart"/>'):kind==='nested'?raw.replace('<w:tcPr>','<w:tbl/><w:tcPr>'):kind==='width'?raw.replace('w:w="2880"','w:w="100"'):kind==='revision'?raw.replace('</w:tcPr>','<w:tcPrChange/></w:tcPr>'):raw.replace('<w:tcPr>','<w:tcPr><w:gridSpan w:val="1junk"/>'));const loaded=await Document.open(p.toBytes()),before=loaded.package.toBytes();expect(()=>loaded.tables[0]!.mergeColumnCells(1,0,2)).toThrow();expect(loaded.package.toBytes()).toEqual(before);}
});

test('protected external settings and stale document handles refuse before writing vertical flags',async()=>{
 for(const mode of ['protected','external','stale']){const d=await fixture(),p=pkg(d),t=d.tables[0]!;if(mode==='protected'){addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');}if(mode==='external')addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','https://example.invalid/settings',{external:true});if(mode==='stale')p.set(p.mainPart(),xml(d).replace('Before','Changed'));const before=d.package.toBytes();expect(()=>t.mergeColumnCells(1,0,2)).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('faults after write and during serialization roll back all markers and preserve held handles',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=await fixture(),p=pkg(d),before=p.toBytes(),t=d.tables[0]!,cell=t.cell(0,1),paragraph=d.paragraphs[0]!,original=p[stage].bind(p);if(stage==='set')p.set=(n,v)=>{original(n as never,v as never);throw Error('injected write');};else p.toBytes=()=>{throw Error('injected serialization');};try{expect(()=>t.mergeColumnCells(1,0,2)).toThrow('injected');}finally{if(stage==='set')p.set=original as OpcPackage['set'];else p.toBytes=original as OpcPackage['toBytes'];}expect(p.toBytes()).toEqual(before);expect(t.rows).toBe(3);expect(cell.text).toBe('Keep & <text>');expect(paragraph.text).toBe('Before');}
});

test('UTF8 BOM and aliased UTF16 LE and BE retain exact sibling XML and flag namespaces after reopen',async()=>{
 for(const encoding of ['bom','le','be']){const d=await fixture(),p=pkg(d);let source=xml(d);if(encoding==='bom')p.set(p.mainPart(),new Uint8Array([239,187,191,...new TextEncoder().encode(source)]));else{source=source.replaceAll('w:','q:').replaceAll('xmlns:w=','xmlns:q=').replace('UTF-8','UTF-16').replace('<q:document ','<q:document xmlns:w="urn:foreign" ');const b=new Uint8Array(2+source.length*2);b.set(encoding==='le'?[255,254]:[254,255]);const view=new DataView(b.buffer);for(let i=0;i<source.length;i++)view.setUint16(2+i*2,source.charCodeAt(i),encoding==='le');p.set(p.mainPart(),b);}const loaded=await Document.open(p.toBytes()),before=loaded.package.parts;loaded.tables[0]!.mergeColumnCells(1,0,2);const read=await Document.open(await loaded.save());expect([...read.package.get('word/document.xml')!.slice(0,encoding==='bom'?3:2)]).toEqual(encoding==='bom'?[239,187,191]:encoding==='le'?[255,254]:[254,255]);expect(withoutFlags(xml(read))).toBe(source);expect(elements(parseXml(xml(read)),'vMerge',W).map(n=>attribute(n,'val',W))).toEqual(['restart','continue','continue']);for(const[n,b]of before)if(n!=='word/document.xml')expect(read.package.get(n)).toEqual(b);}
});

test('post-merge editable coordinates row changes and repeat merges refuse while unrelated cells remain readable',async()=>{
 const d=await fixture(),t=d.tables[0]!.mergeColumnCells(1,0,2),before=d.package.toBytes();for(const action of [()=>t.cell(0,1),()=>t.tryCell(1,1),()=>t.cell(2,1),()=>t.insertRow(0),()=>t.mergeColumnCells(0,0,2),()=>t.mergeRowCells(0,0,1)]){expect(action).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(t.cell(2,2).text).toBe('Tail');expect(t.tryCell(3,1)).toBeNull();
});

test('vertical owner formatting follows schema order and complete table admission refuses section or lexical barriers',async()=>{
 const d=await fixture();d.tables[0]!.cell(0,1).setProperties({topBorder:{style:'single',size:8,color:'123456'},textDirection:'tbRl'});const source=xml(d);d.tables[0]!.mergeColumnCells(1,0,2);expect(withoutFlags(xml(d))).toBe(source);expect(cells(xml(d))[1]!.children[0]!.children.map(n=>n.localName)).toEqual(['tcW','vMerge','tcBorders','shd','textDirection','vAlign']);
 for(const kind of ['section','lexical','width-decoration']){const d=await fixture(),p=pkg(d);replaceCell(p,8,raw=>kind==='section'?raw.replace(/<w:p\b[^>]*>/,open=>open+'<w:pPr><w:sectPr/></w:pPr>'):kind==='lexical'?raw.replace('<w:tcPr>','<!--retain--><w:tcPr>'):raw.replace('w:type="dxa"','w:type="dxa" retain="yes"'));const loaded=await Document.open(p.toBytes()),before=loaded.package.toBytes();expect(()=>loaded.tables[0]!.mergeColumnCells(1,0,2)).toThrow();expect(loaded.package.toBytes()).toEqual(before);}
});

test('vertical authoring leaves all unselected tables and package members unchanged',async()=>{
 const d=await fixture();d.addTable(1,1);d.tables[1]!.cell(0,0).text='Other table';const source=xml(d),before=d.package.parts;d.tables[0]!.mergeColumnCells(1,0,2);expect(withoutFlags(xml(d))).toBe(source);expect(d.tables[1]!.cell(0,0).text).toBe('Other table');expect([...d.package.parts.keys()]).toEqual([...before.keys()]);for(const[n,b]of before)if(n!=='word/document.xml')expect(d.package.get(n)).toEqual(b);
});

test('vertical native authoring does not select incompatible shared getter or horizontal outcomes',async()=>{
 const {inventoryFeatures}=await import('../../scripts/gherkin.ts'),inv=await inventoryFeatures(process.cwd()),s=inv.features.flatMap(f=>f.scenarios).find(s=>s.scenarioId==='@id-docx-go-table-merge-properties')!;expect(s.lifecycle).toBe('planned');expect(s.cases).toHaveLength(1);inv.counts.cases=inv.coverage!.shared.cases;expect(inv.counts.cases.implemented).toBe(732);expect(inv.counts.cases.planned).toBe(59);
});
