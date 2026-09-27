import {expect,test} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {tmpdir} from 'node:os';import {join} from 'node:path';
import {Document,type CellPropertiesPatch} from '../../src/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,attribute,applyEdits} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(properties?:string){const d=Document.create();d.addTable(2,2);d.tables[0]!.cell(0,0).text='Alpha';d.tables[0]!.cell(0,1).text='Untouched';if(properties===undefined)return d;const p=await OpcPackage.open(d.package.toBytes()),xml=p.text(p.mainPart()),cell=elements(parseXml(xml),'tc',W)[0]!,pr=cell.children.find(n=>n.localName==='tcPr')!;p.set(p.mainPart(),applyEdits(xml,[{start:pr.start,end:pr.end,value:properties}]));return Document.open(p.toBytes());}
const props={widthTwips:2400,verticalAlign:'center',textDirection:'tbRl',shading:'FFFF00',topBorder:{style:'single',size:8,color:'000000'}} as const;

test('direct cell properties survive reopen without changing text, grid, sibling cells or other package parts',async()=>{
 const d=await fixture(),before=d.package.parts,table=d.tables[0]!,cell=table.cell(0,0),xmlBefore=new TextDecoder().decode(before.get('word/document.xml'));
 const sibling=elements(parseXml(xmlBefore),'tc',W)[1]!,siblingXml=xmlBefore.slice(sibling.start,sibling.end);
 expect(cell.directProperties()).toEqual({widthTwips:4320,verticalAlign:null,textDirection:null,shading:null,topBorder:null});
 expect(cell.setProperties(props)).toEqual({changed:1});expect(()=>cell.directProperties()).toThrow(expect.objectContaining({code:'docx-stale-table-cell'}));
 expect(table.rows).toBe(2);expect(table.columns).toBe(2);expect(table.cell(0,0).text).toBe('Alpha');
 const reopened=await Document.open(d.package.toBytes());expect(reopened.tables[0]!.cell(0,0).directProperties()).toEqual(props);
 const xml=new TextDecoder().decode(reopened.package.get('word/document.xml'));expect(xml).toContain(siblingXml);
 const parsed=parseXml(xml);expect(elements(parsed,'tcPr',W)[0]!.children.map(n=>n.localName)).toEqual(['tcW','tcBorders','shd','textDirection','vAlign']);
 const width=elements(parsed,'tcW',W)[0]!;expect(attribute(width,'w',W)).toBe('2400');expect(attribute(width,'type',W)).toBe('dxa');
 for(const[n,b]of before)if(n!=='word/document.xml')expect(reopened.package.get(n)).toEqual(b);
});

test('null removes direct properties, same-state preserves archive and returned border data is detached',async()=>{
 const d=await fixture();d.tables[0]!.cell(0,0).setProperties(props);const cell=d.tables[0]!.cell(0,0),before=d.package.toBytes();
 expect(cell.setProperties(props)).toEqual({changed:0});expect(d.package.toBytes()).toEqual(before);expect(cell.text).toBe('Alpha');
 const value=cell.directProperties();value.topBorder!.size=16;expect(cell.directProperties().topBorder!.size).toBe(8);
 cell.setProperties({widthTwips:null,verticalAlign:null,textDirection:null,shading:null,topBorder:null});expect(d.tables[0]!.cell(0,0).directProperties()).toEqual({widthTwips:null,verticalAlign:null,textDirection:null,shading:null,topBorder:null});
 const after=d.package.toBytes();expect(d.tables[0]!.cell(0,0).setProperties({topBorder:null})).toEqual({changed:0});expect(d.package.toBytes()).toEqual(after);
});

test('editing top border preserves other border fragments and unrelated cell properties exactly',async()=>{
 const side='<w:bottom w:val = \'dashed\' w:sz="12" w:color="123456"/>';
 const pr=`<w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:tcBorders>${side}</w:tcBorders><w:noWrap/><w:tcMar><w:top w:w="40" w:type="dxa"/></w:tcMar></w:tcPr>`;
 const d=await fixture(pr);d.tables[0]!.cell(0,0).setProperties({topBorder:props.topBorder,shading:'aabbcc'});
 let xml=new TextDecoder().decode(d.package.get('word/document.xml'));expect(xml).toContain(side);expect(xml).toContain('<w:noWrap/>');expect(xml).toContain('<w:tcMar><w:top w:w="40" w:type="dxa"/></w:tcMar>');
 d.tables[0]!.cell(0,0).setProperties({topBorder:null});xml=new TextDecoder().decode(d.package.get('word/document.xml'));expect(xml).toContain(side);expect(d.tables[0]!.cell(0,0).directProperties().topBorder).toBeNull();
});

test('invalid and executable nested patches refuse without invoking accessors or mutation',async()=>{
 const d=await fixture(),cell=d.tables[0]!.cell(0,0),before=d.package.toBytes();let called=false;
 const bad:unknown[]=[null,{}, {widthTwips:-1},{widthTwips:31681},{widthTwips:1.5},{verticalAlign:'both'},{textDirection:'sideways'},{shading:'#FFFF00'},{shading:'yellow'},{topBorder:{style:'double',size:8,color:'000000'}},{topBorder:{style:'single',size:1,color:'000000'}},{topBorder:{style:'single',size:97,color:'000000'}},{topBorder:{style:'single',size:8,color:'bad'}},{topBorder:{style:'single',size:8,color:'000000',extra:1}},{shading:'FFFF00',widthTwips:'bad'}];
 bad.push({topBorder:Object.defineProperty({size:8,color:'000000'},'style',{enumerable:true,get(){called=true;return 'single';}})});
 for(const patch of bad){expect(()=>cell.setProperties(patch as CellPropertiesPatch)).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(called).toBe(false);
});

test('selected percentage widths, themed shading and decorated top borders refuse rather than dropping metadata',async()=>{
 const cases:[string,CellPropertiesPatch][]=[
 ['<w:tcW w:w="5000" w:type="pct"/>',{widthTwips:2400}],
 ['<w:shd w:val="clear" w:fill="FFFF00" w:themeFill="accent1"/>',{shading:'FFFF00'}],
 ['<w:shd w:val="pct20" w:fill="FFFF00"/>',{shading:null}],
 ['<w:tcBorders><w:top w:val="single" w:sz="8" w:color="000000" w:shadow="1"/></w:tcBorders>',{topBorder:null}],
 ['<w:tcBorders><w:top w:val="single" w:sz="8" w:color="000000"/><w:top w:val="single" w:sz="8" w:color="000000"/></w:tcBorders>',{topBorder:null}],
 ['<w:vAlign val="center"/>',{verticalAlign:'center'}],['<w:textDirection w:val="unknown"/>',{textDirection:null}],
 ];
 for(const[leaf,patch]of cases){const d=await fixture('<w:tcPr>'+leaf+'</w:tcPr>'),before=d.package.toBytes();expect(()=>d.tables[0]!.cell(0,0).setProperties(patch)).toThrow();expect(()=>d.tables[0]!.cell(0,0).directProperties()).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('unknown, duplicate, revised or lexical cell metadata refuses before a same-value request',async()=>{
 for(const pr of ['<w:tcPr><w:mystery/></w:tcPr>','<w:tcPr><w:tcPrChange/></w:tcPr>','<w:tcPr><w:vAlign w:val="center"/><w:shd w:val="clear" w:fill="FFFF00"/></w:tcPr>','<w:tcPr><!--barrier--><w:shd w:val="clear" w:fill="FFFF00"/></w:tcPr>','<w:tcPr/><w:tcPr/>']){
  const d=await fixture(pr),before=d.package.toBytes();expect(()=>d.tables[0]!.cell(0,0).setProperties({shading:'FFFF00'})).toThrow();expect(d.package.toBytes()).toEqual(before);
 }
});

test('protection and external document changes refuse without package mutation',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');
 const locked=await Document.open(pkg.toBytes()),before=locked.package.toBytes();expect(()=>locked.tables[0]!.cell(0,0).setProperties({shading:'FFFF00'})).toThrow();expect(locked.package.toBytes()).toEqual(before);
 const held=d.tables[0]!.cell(0,0);d.package.setPart('word/document.xml',new TextEncoder().encode(new TextDecoder().decode(d.package.get('word/document.xml')).replace('Alpha','Beta')));const changed=d.package.toBytes();expect(()=>held.setProperties({shading:'FFFF00'})).toThrow();expect(()=>held.directProperties()).toThrow();expect(d.package.toBytes()).toEqual(changed);
});

test('staged write and serialization failures retain the cell handle and bytes',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=await fixture(),cell=d.tables[0]!.cell(0,0),before=d.package.toBytes(),pkg=(d as unknown as {opcPackage:OpcPackage}).opcPackage,original=pkg[stage].bind(pkg);
  if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialize');};
  try{expect(()=>cell.setProperties(props)).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}
  expect(d.package.toBytes()).toEqual(before);expect(cell.text).toBe('Alpha');expect(cell.directProperties().shading).toBeNull();
 }
});

test('alias namespaces and UTF-16 path save preserve values and unrelated payloads',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()).replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('UTF-8','UTF-16');const bytes=new Uint8Array(2+text.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);
 const doc=await Document.open(pkg.toBytes());doc.tables[0]!.cell(0,0).setProperties(props);const root=await mkdtemp(join(tmpdir(),'cell-properties-'));
 try{const path=join(root,'table.docx');await doc.save(path);const reopened=await Document.open(path);expect(reopened.tables[0]!.cell(0,0).directProperties()).toEqual(props);expect([...reopened.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(reopened.package.get(n)).toEqual(pkg.get(n));}finally{await rm(root,{recursive:true,force:true});}
});

test('missing or self-closing cell property containers accept ordered additions without changing children',async()=>{
 for(const pr of ['', '<w:tcPr/>','<w:tcPr><w:tcBorders/></w:tcPr>']){const d=await fixture(pr);d.tables[0]!.cell(0,0).setProperties(props);expect(d.tables[0]!.cell(0,0).directProperties()).toEqual(props);expect(d.tables[0]!.cell(0,0).text).toBe('Alpha');expect((await Document.open(d.package.toBytes())).tables[0]!.cell(0,1).text).toBe('Untouched');}
});

test('misordered owner properties and nested cell-property revisions refuse atomically',async()=>{
 for(const change of [
  (xml:string)=>xml.replace('<w:tr>','<w:tr><w:trPr/>').replace('</w:tr>','<w:trPr/></w:tr>'),
  (xml:string)=>xml.replace('</w:tbl>','<w:tblPr/></w:tbl>'),
  (xml:string)=>xml.replace(/<w:tblGrid>[^]*?<\/w:tblGrid>/,''),
  (xml:string)=>xml.replace('</w:tcPr>','<w:tcMar><w:tcPrChange/></w:tcMar></w:tcPr>'),
 ]){const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),change(pkg.text(pkg.mainPart())));const modified=await Document.open(pkg.toBytes()),before=modified.package.toBytes();expect(()=>modified.tables[0]!.cell(0,0).setProperties({shading:'FFFF00'})).toThrow();expect(modified.package.toBytes()).toEqual(before);}
});

test('merged or nested cells refuse property writes and no-op still checks protection',async()=>{
 const d=await fixture('<w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>');expect(()=>d.tables[0]!.cell(0,0)).toThrow(expect.objectContaining({code:'docx-table-merged-cell'}));
 const normal=await fixture(),pkg=await OpcPackage.open(normal.package.toBytes());pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('</w:tc>','<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc>'));const nested=await Document.open(pkg.toBytes()),nb=nested.package.toBytes();expect(()=>nested.tables[0]!.cell(0,0).setProperties({shading:'FFFF00'})).toThrow();expect(nested.package.toBytes()).toEqual(nb);
 const protectedPkg=await OpcPackage.open(normal.package.toBytes());addPart(protectedPkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(protectedPkg,protectedPkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(protectedPkg.toBytes());expect(()=>locked.tables[0]!.cell(0,0).setProperties({widthTwips:4320})).toThrow();
});

test('two shared direct cell cases execute and altered readback predicates fail',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts');const {bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/cell-properties.ts');const path='workflows/docx/document-model.feature',text=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv=(s:string)=>({root:'.',features:[selectSharedScenarios(path,s,scenarioIds)],counts:{features:count(1),scenarios:count(2),cases:count(2),steps:count(8)}});
 const good=await executeAcceptance(inv(text),bindings,'cell-properties-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(2);
 for(const[from,to]of [['its shading getter equals FFFF00','its shading getter equals 000000'],['width equals 2400, width type dxa, alignment center and direction tbRl','width equals 999, width type pct, alignment top and direction lrTb']]){const bad=await executeAcceptance(inv(text.replace(from!,to!)),bindings,'cell-properties-control');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 const corrupt=bindings.map(b=>b.pattern.test('a top border with single style, size eight and colour 000000 is assigned')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);(c.state as {document:Document}).document.tables[0]!.cell(0,0).setProperties({topBorder:null});}}:b);
 const bad=await executeAcceptance(inv(text),corrupt,'cell-border-control');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);
});

test('bounded width, border size and all supported direction/alignment values survive direct readback',async()=>{
 const d=await fixture();
 for(const widthTwips of [0,31680])for(const size of [2,96]){d.tables[0]!.cell(0,0).setProperties({widthTwips,topBorder:{style:'single',size,color:'auto'}});const v=d.tables[0]!.cell(0,0).directProperties();expect(v.widthTwips).toBe(widthTwips);expect(v.topBorder).toEqual({style:'single',size,color:'auto'});}
 for(const verticalAlign of ['top','center','bottom'] as const)for(const textDirection of ['lrTb','tbRl','btLr'] as const){d.tables[0]!.cell(0,0).setProperties({verticalAlign,textDirection,shading:'auto'});const v=d.tables[0]!.cell(0,0).directProperties();expect(v.verticalAlign).toBe(verticalAlign);expect(v.textDirection).toBe(textDirection);expect(v.shading).toBe('auto');}
 const before=d.package.toBytes();expect(()=>d.tables[0]!.cell(0,0).setProperties({shading:'FFFF00',topBorder:{style:'single',size:97,color:'000000'}})).toThrow();expect(d.package.toBytes()).toEqual(before);
});

test('colour validation refuses trailing line endings rather than XML-normalising a requested value',async()=>{
 const d=await fixture(),before=d.package.toBytes();
 for(const color of ['FFFF00\n','auto\r','000000\r\n'])for(const patch of [{shading:color},{topBorder:{style:'single' as const,size:8,color}}]){
  expect(()=>d.tables[0]!.cell(0,0).setProperties(patch)).toThrow();expect(d.package.toBytes()).toEqual(before);
 }
});
