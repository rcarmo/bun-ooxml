import {test,expect} from 'bun:test';import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';import {addPart,addRelationship} from '../../src/opc/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const pkg=(d:Document)=>(d as any).opcPackage as OpcPackage;
async function sample(){const d=Document.create();d.addParagraph('Statement for <Customer>').setProperties({outlineLevel:0});d.addParagraph('Use <Project> and [TBD].');d.addTable(2,2);const t=d.tables[0]!;t.cell(0,0).text='Role';t.cell(0,1).text='Hours';t.cell(1,0).text='<Role>';t.cell(1,1).text='40';d.addParagraph('Delivery').setProperties({outlineLevel:1});d.addParagraph('No placeholder');const p=await OpcPackage.open(d.package.toBytes());addPart(p,'customXml/opaque.bin',new Uint8Array([0,255,42]),'application/octet-stream');return Document.open(p.toBytes());}

test('template inspection returns concrete sections placeholders locations and table text without mutating archive',async()=>{
 const d=await sample(),before=d.package.toBytes(),t=d.tables[0]!,held=d.paragraphs[0]!,r=d.inspectTemplate();expect(r.scope).toBe('main-body-and-tables');expect(r.counts).toEqual({paragraphs:8,sections:2,tables:1,placeholders:3});expect(r.paragraphs.map(p=>p.text)).toEqual(['Statement for <Customer>','Use <Project> and [TBD].','Role','Hours','<Role>','40','Delivery','No placeholder']);
 expect(r.sections).toEqual([{paragraphIndex:0,bodyIndex:0,text:'Statement for <Customer>',level:1},{paragraphIndex:6,bodyIndex:3,text:'Delivery',level:2}]);expect(r.tables).toEqual([{index:0,bodyIndex:2,rows:2,columns:2,cells:[['Role','Hours'],['<Role>','40']]}]);expect(r.paragraphs[4]!.table).toEqual({index:0,row:1,column:0,paragraph:0});expect(r.placeholders).toEqual([{paragraphIndex:0,bodyIndex:0,start:14,end:24,text:'<Customer>',name:'Customer'},{paragraphIndex:1,bodyIndex:1,start:4,end:13,text:'<Project>',name:'Project'},{paragraphIndex:4,bodyIndex:2,start:0,end:6,text:'<Role>',name:'Role'}]);expect(d.package.toBytes()).toEqual(before);expect(held.text).toBe('Statement for <Customer>');expect(t.cell(1,1).text).toBe('40');
});

test('inspection persists through path reopen and results are detached and deeply immutable',async()=>{
 const d=await sample(),before=d.package.parts,r=d.inspectTemplate(),dir=await mkdtemp(join(tmpdir(),'template-inspect-'));try{const path=join(dir,'template.docx');await d.save(path);const read=await Document.open(path);expect(read.inspectTemplate()).toEqual(r);for(const[n,b]of before)expect(read.package.get(n)).toEqual(b);}finally{await rm(dir,{recursive:true,force:true});}
 for(const value of [r,r.counts,r.paragraphs,r.paragraphs[4],r.paragraphs[4]!.table,r.sections,r.sections[0],r.tables,r.tables[0],r.tables[0]!.cells,r.tables[0]!.cells[0],r.placeholders,r.placeholders[0]])expect(Object.isFrozen(value)).toBe(true);
 d.paragraphs.at(-1)!.setText('Changed <New>');expect(r.paragraphs.at(-1)!.text).toBe('No placeholder');expect(r.counts.placeholders).toBe(3);expect(d.inspectTemplate().counts.placeholders).toBe(4);
});

test('plain and empty documents return useful empty inventories without inventing classifications',()=>{
 const d=Document.create();expect(d.inspectTemplate().counts).toEqual({paragraphs:0,sections:0,tables:0,placeholders:0});d.addParagraphStyle('Heading1',{name:'Heading 1'});d.addParagraph('Blue <Customer>',{style:'Heading1'}).setRunFormatting({color:'0000FF'});const r=d.inspectTemplate();expect(r.sections).toEqual([]);expect(r.paragraphs[0]!.styleId).toBe('Heading1');expect(r.paragraphs[0]!.outlineLevel).toBeNull();expect(r.placeholders.map(p=>p.name)).toEqual(['Customer']);expect(Object.keys(r).sort()).toEqual(['counts','paragraphs','placeholders','scope','sections','tables']);
});

test('placeholders span plain runs but never paragraphs and retain Unicode UTF16 offsets and occurrence order',()=>{
 const d=Document.create();d.addParagraph('😀 <Cus').appendRun('tomer> <Customer> <<nested>> <> < > [TBD]');d.addParagraph('<line\nbreak>');d.addParagraph('<unfinished');d.addParagraph('ending>');const r=d.inspectTemplate();expect(r.placeholders.map(p=>({start:p.start,end:p.end,name:p.name}))).toEqual([{start:3,end:13,name:'Customer'},{start:14,end:24,name:'Customer'}]);expect(d.inspectBodyMap().placeholders.map(p=>({start:p.start,end:p.end,name:p.name}))).toEqual(r.placeholders.map(p=>({start:p.start,end:p.end,name:p.name})));
});

test('table cell paragraphs have distinct locations and cannot create cross-paragraph placeholders or sections',()=>{
 const d=Document.create();d.addTable(1,1);d.tables[0]!.cell(0,0).text='<A>\n<B>\n<split\nend>';const r=d.inspectTemplate();expect(r.tables[0]!.cells).toEqual([['<A>\n<B>\n<split\nend>']]);expect(r.paragraphs.map(p=>p.table!.paragraph)).toEqual([0,1,2,3]);expect(r.placeholders.map(p=>[p.paragraphIndex,p.name])).toEqual([[0,'A'],[1,'B']]);expect(r.sections).toEqual([]);
});

test('unsupported late body and cell content refuses the whole inspection with unchanged bytes',async()=>{
 for(const kind of ['field','hyperlink','sdt','nested','merge','bad-grid','bad-outline']){const d=await sample(),p=pkg(d);let x=p.text(p.mainPart());if(kind==='field')x=x.replace('No placeholder','No placeholder</w:t><w:instrText>DATE</w:instrText><w:t>');if(kind==='hyperlink')x=x.replace('</w:body>','<w:p><w:hyperlink><w:r><w:t>Hidden</w:t></w:r></w:hyperlink></w:p></w:body>');if(kind==='sdt')x=x.replace('<w:body>','<w:body><w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt>');if(kind==='nested')x=x.replace('<w:tcPr>','<w:tbl/><w:tcPr>');if(kind==='merge')x=x.replace('<w:tcPr>','<w:tcPr><w:vMerge/>');if(kind==='bad-grid')x=x.replace(/<w:gridCol[^>]*\/>/,'');if(kind==='bad-outline')x=x.replace('w:val="0"','w:val="invalid"');p.set(p.mainPart(),x);const loaded=await Document.open(p.toBytes()),before=loaded.package.toBytes();expect(()=>loaded.inspectTemplate()).toThrow();expect(loaded.package.toBytes()).toEqual(before);}
});

test('stale external document edits refuse and no header story or macro text is silently included',async()=>{
 const d=await sample(),p=pkg(d);addPart(p,'word/header1.xml',`<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>&lt;Header&gt;</w:t></w:r></w:p></w:hdr>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header','header1.xml');const before=d.package.toBytes();expect(d.inspectTemplate().placeholders.map(p=>p.name)).toEqual(['Customer','Project','Role']);expect(d.package.toBytes()).toEqual(before);p.set(p.mainPart(),p.text(p.mainPart()).replace('No placeholder','External'));const changed=p.toBytes();expect(()=>d.inspectTemplate()).toThrow(expect.objectContaining({code:'docx-stale-document'}));expect(d.package.toBytes()).toEqual(changed);
});

test('aliased UTF16BE inspection preserves exact source bytes and namespace meaning',async()=>{
 const d=await sample(),expected=d.inspectTemplate(),p=pkg(d);const text=p.text(p.mainPart()).replaceAll('w:','q:').replaceAll('xmlns:w=','xmlns:q=').replace('UTF-8','UTF-16').replace('<q:document ','<q:document xmlns:w="urn:foreign" '),bytes=new Uint8Array(2+text.length*2);bytes.set([254,255]);const view=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),false);p.set(p.mainPart(),bytes);const read=await Document.open(p.toBytes()),before=read.package.toBytes();expect(read.inspectTemplate()).toEqual(expected);expect(read.package.toBytes()).toEqual(before);
});

test('global placeholder and paragraph bounds refuse instead of returning a truncated successful report',async()=>{
 const d=Document.create();d.addParagraph('<A>'.repeat(10001));const before=d.package.toBytes();expect(()=>d.inspectTemplate()).toThrow(expect.objectContaining({code:'docx-template-limit'}));expect(d.package.toBytes()).toEqual(before);
 const base=Document.create(),p=pkg(base);p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:body>','<w:body>'+'<w:p/>'.repeat(10001)));const loaded=await Document.open(p.toBytes()),b=loaded.package.toBytes();expect(()=>loaded.inspectTemplate()).toThrow(expect.objectContaining({code:'docx-template-limit'}));expect(loaded.package.toBytes()).toEqual(b);
});

test('grid lexical decorations wrong namespaces and late property ambiguity refuse read-only',async()=>{
 for(const change of [(x:string)=>x.replace('<w:tblGrid>','<w:tblGrid><!--keep-->'),(x:string)=>x.replace('<w:gridCol ','<w:gridCol other="x" '),(x:string)=>x.replace('xmlns:w="'+W+'"','xmlns:w="urn:foreign"'),(x:string)=>x.replace(/<w:pPr\b[^>]*>/,open=>open+'<w:outlineLvl w:val="1"/>')]){const d=await sample(),p=pkg(d);const source=p.text(p.mainPart()),changed=change(source);expect(changed).not.toBe(source);p.set(p.mainPart(),changed);let read:Document;try{read=await Document.open(p.toBytes());}catch(e){expect(e).toBeInstanceOf(Error);continue;}const before=read.package.toBytes();expect(()=>read.inspectTemplate()).toThrow();expect(read.package.toBytes()).toEqual(before);}
});

test('table count and aggregate placeholder limits apply across cells with no successful truncation',async()=>{
 const d=Document.create();d.addTable(1,2);d.tables[0]!.cell(0,0).text='<A>'.repeat(5001);d.tables[0]!.cell(0,1).text='<B>'.repeat(5000);const before=d.package.toBytes();expect(()=>d.inspectTemplate()).toThrow(expect.objectContaining({code:'docx-template-limit'}));expect(d.package.toBytes()).toEqual(before);
 const base=Document.create();base.addTable(1,1);const p=pkg(base),source=p.text(p.mainPart()),table=source.match(/<w:tbl\b[^]*?<\/w:tbl>/)![0];p.set(p.mainPart(),source.replace(table,table.repeat(1001)));const read=await Document.open(p.toBytes()),bytes=read.package.toBytes();expect(()=>read.inspectTemplate()).toThrow(expect.objectContaining({code:'docx-template-limit'}));expect(read.package.toBytes()).toEqual(bytes);
});

test('read-only inspection permits protected input and leaves all template response/cache scenarios planned',async()=>{
 const d=await sample(),p=pkg(d);addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const before=d.package.toBytes();expect(d.inspectTemplate().counts.placeholders).toBe(3);expect(d.package.toBytes()).toEqual(before);
 const {inventoryFeatures}=await import('../../scripts/gherkin.ts'),inv=await inventoryFeatures(process.cwd()),templates=inv.features.filter(f=>/\/template-(analysis|cache)\.feature$/.test(f.path));expect(templates).toHaveLength(2);expect(templates.flatMap(f=>f.scenarios).filter(s=>!s.scenarioId.startsWith('@id-docx-template-inventory-'))).toHaveLength(13);expect(templates.flatMap(f=>f.scenarios.map(s=>({s,f}))).filter(({s})=>!s.scenarioId.startsWith('@id-docx-template-inventory-')).every(({s,f})=>(s.lifecycle??f.lifecycle)==='planned')).toBe(true);expect(inv.counts.cases.implemented).toBe(726);expect(inv.counts.cases.planned).toBe(46);
});

test('row offsets and mismatched or absent cell widths refuse rather than reporting false rectangular coordinates',async()=>{
 for(const defect of ['before','after','late-before','width','missing-width']){const d=await sample(),p=pkg(d);let source=p.text(p.mainPart());if(defect==='before'||defect==='after')source=source.replace('<w:tr>',`<w:tr><w:trPr><w:${defect==='before'?'gridBefore':'gridAfter'} w:val="1"/></w:trPr>`);else if(defect==='late-before'){const at=source.lastIndexOf('<w:tr>');source=source.slice(0,at)+source.slice(at).replace('<w:tr>','<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>');}else if(defect==='width')source=source.replace('w:w="4320" w:type="dxa"','w:w="100" w:type="dxa"');else source=source.replace(/<w:tcW[^>]*\/>/,'');expect(source).not.toBe(p.text(p.mainPart()));p.set(p.mainPart(),source);const read=await Document.open(p.toBytes()),before=read.package.toBytes();expect(()=>read.inspectTemplate()).toThrow(expect.objectContaining({code:'docx-template-unsupported'}));expect(read.package.toBytes()).toEqual(before);}
});

test('explicit zero row offsets retain correct positions and invalid offset spellings refuse',async()=>{
 const d=await sample(),expected=d.inspectTemplate(),p=pkg(d);p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:tr>','<w:tr><w:trPr><w:gridBefore w:val="0"/><w:gridAfter w:val="00"/></w:trPr>'));const read=await Document.open(p.toBytes()),before=read.package.toBytes();expect(read.inspectTemplate()).toEqual(expected);expect(read.package.toBytes()).toEqual(before);
 for(const value of ['-1','0junk','', '1.5']){const d=await sample(),p=pkg(d);p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:tr>',`<w:tr><w:trPr><w:gridBefore w:val="${value}"/></w:trPr>`));const read=await Document.open(p.toBytes()),before=read.package.toBytes();expect(()=>read.inspectTemplate()).toThrow();expect(read.package.toBytes()).toEqual(before);}
});
