import {test,expect} from 'bun:test';
import {Document,OpcPackage} from '../../src/index.ts';
import {addPart} from '../../src/opc/index.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {tmpdir} from 'node:os';import {join} from 'node:path';
async function fixture(){const d=Document.create();d.addTable(3,3);const t=d.tables[0]!;for(let r=0;r<3;r++)for(let c=0;c<3;c++)t.cell(r,c).text=`${r},${c}`;const p=await OpcPackage.open(d.package.toBytes());addPart(p,'customXml/opaque.bin',new Uint8Array([0,255,42]),'application/octet-stream');return Document.open(p.toBytes());}

test('nullable table access returns nine usable cells and null at each integer boundary without mutation',async()=>{
 const d=await fixture(),t=d.tables[0]!,before=d.package.toBytes();for(let r=0;r<3;r++)for(let c=0;c<3;c++){const cell=t.tryCell(r,c);expect(cell).not.toBeNull();expect(cell!.text).toBe(`${r},${c}`);expect(cell!.directProperties()).toEqual(t.cell(r,c).directProperties());}
 for(const[r,c]of [[-1,0],[0,-1],[3,0],[0,3],[-1,-1],[3,3],[Number.MAX_SAFE_INTEGER,0],[0,-Number.MAX_SAFE_INTEGER]]){expect(t.tryCell(r!,c!)).toBeNull();expect(()=>t.cell(r!,c!)).toThrow(RangeError);}expect(d.package.toBytes()).toEqual(before);
});

test('nullable access refuses invalid numeric coordinates without coercion even if the other axis is outside',async()=>{
 const d=await fixture(),t=d.tables[0]!,before=d.package.toBytes();let called=false;for(const value of [NaN,Infinity,-Infinity,0.5,-0.5,Number.MAX_SAFE_INTEGER+1,'0',null,undefined,1n,{valueOf(){called=true;return 0;}}]){expect(()=>t.tryCell(value as any,0)).toThrow(RangeError);expect(()=>t.tryCell(0,value as any)).toThrow(RangeError);expect(()=>t.tryCell(-1,value as any)).toThrow(RangeError);expect(d.package.toBytes()).toEqual(before);}expect(called).toBe(false);expect(t.tryCell(-0,-0)!.text).toBe('0,0');
});

test('nullable cells edit through the existing checked path and saved tables reopen with all original coordinates',async()=>{
 const d=await fixture(),t=d.tables[0]!,held=t.tryCell(1,1)!,before=d.package.parts;held.text='Updated & <cell>';expect(()=>held.text).toThrow(expect.objectContaining({code:'docx-stale-table-cell'}));expect(t.tryCell(1,1)!.text).toBe('Updated & <cell>');
 const dir=await mkdtemp(join(tmpdir(),'nullable-cell-'));try{const path=join(dir,'table.docx');await d.save(path);const reopened=await Document.open(path);expect(reopened.tables).toHaveLength(1);expect(reopened.tables[0]!.rows).toBe(3);expect(reopened.tables[0]!.columns).toBe(3);for(let r=0;r<3;r++)for(let c=0;c<3;c++)expect(reopened.tables[0]!.tryCell(r,c)!.text).toBe(r===1&&c===1?'Updated & <cell>':`${r},${c}`);expect(reopened.tables[0]!.tryCell(3,0)).toBeNull();expect([...reopened.package.parts.keys()]).toEqual([...before.keys()]);for(const[n,b]of before)if(n!=='word/document.xml')expect(reopened.package.get(n)).toEqual(b);}finally{await rm(dir,{recursive:true,force:true});}
});

test('stale structural handles refuse both in-range and out-of-range nullable lookup',async()=>{
 const d=await fixture(),t=d.tables[0]!;t.insertRow(0);const before=d.package.toBytes();for(const[r,c]of [[0,0],[-1,0],[3,3]])expect(()=>t.tryCell(r!,c!)).toThrow(expect.objectContaining({code:'docx-stale-table'}));expect(d.package.toBytes()).toEqual(before);
});

test('externally changed table bytes cannot masquerade as a missing cell',async()=>{
 const d=await fixture(),t=d.tables[0]!,source=new TextDecoder().decode(d.package.get('word/document.xml'));d.package.setPart('word/document.xml',new TextEncoder().encode(source.replace('0,0','External')));const before=d.package.toBytes();expect(()=>t.tryCell(-1,0)).toThrow(expect.objectContaining({code:'docx-stale-table'}));expect(()=>t.tryCell(0,0)).toThrow(expect.objectContaining({code:'docx-stale-table'}));expect(d.package.toBytes()).toEqual(before);
});

test('nullable access does not hide merged cells unsupported grids or unsafe cell text',async()=>{
 for(const kind of ['merged','grid','nested']){const d=await fixture(),p=await OpcPackage.open(d.package.toBytes()),xml=p.text(p.mainPart());p.set(p.mainPart(),kind==='merged'?xml.replace('<w:tcPr>','<w:tcPr><w:vMerge w:val="restart"/>'):kind==='grid'?xml.replace(/<w:gridCol[^>]*\/>/,''):xml.replace('<w:tcPr>','<w:tbl/><w:tcPr>'));const loaded=await Document.open(p.toBytes()),t=loaded.tables[0]!,before=loaded.package.toBytes();if(kind==='merged'){expect(()=>t.tryCell(0,0)).toThrow(expect.objectContaining({code:'docx-table-merged-cell'}));expect(t.tryCell(3,0)).toBeNull();}else if(kind==='grid'){expect(()=>t.tryCell(0,0)).toThrow(expect.objectContaining({code:'docx-table-unsupported'}));expect(()=>t.tryCell(-1,0)).toThrow(expect.objectContaining({code:'docx-table-unsupported'}));}else{const cell=t.tryCell(0,0);expect(cell).not.toBeNull();expect(()=>cell!.text).toThrow(expect.objectContaining({code:'docx-table-cell-unsupported'}));}expect(loaded.package.toBytes()).toEqual(before);}
});

test('shared nullable profile executes one case and detects missing in-range cells or fabricated boundary cells',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/nullable-cell.ts');const path='workflows/docx/tables.feature',text=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv={root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(4)}};
 const good=await executeAcceptance(inv,bindings,'nullable-cell-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);
 const changed=bindings.map(b=>b.pattern.test('its Cell getter is called for all nine coordinates from zero through two')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);(c.nullableCells as unknown[])[4]=null;}}:b);const bad=await executeAcceptance(inv,changed,'nullable-cell-corrupt');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);
 const {Table}=await import('../../src/docx/index.ts'),original=Table.prototype.tryCell;Table.prototype.tryCell=function(r,c){return r<0||c<0||r>=this.rows||c>=this.columns?this.cell(0,0):original.call(this,r,c);};try{const boundary=await executeAcceptance(inv,bindings,'nullable-cell-boundary');expect(boundary.counts.cases.failed).toBe(1);expect(boundary.counts.steps.undefined).toBe(0);}finally{Table.prototype.tryCell=original;}
});
