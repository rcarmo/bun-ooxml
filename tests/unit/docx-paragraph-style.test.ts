import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings,styleDocument,directStyle,W,STYLE_TYPE,STYLE_REL} from '../acceptance/paragraph-style.ts';
import {Document} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';

test('paragraph style selection executes19 saved outcome and refusal cases',async()=>{
 const {sharedScenarios}=await import('../helpers/shared-scenarios.ts');const [f]=await sharedScenarios(["@id-docx-paragraph-style-selection","@id-docx-paragraph-style-refusal"]);if(!f)throw Error('Missing shared profile');const rows=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(rows.length),steps:count(rows.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inv,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(19);
});
test('changed style invalidates handles but no-op preserves them',async()=>{
 const d=await styleDocument('same-style'),p=d.paragraphs[0]!,span=p.find('Styled')[0]!;expect(p.styleId).toBe('Heading1');expect(p.setStyle('Heading1').changed).toBe(0);expect(p.text).toBe('Styled π text');expect(p.setStyle('Body').changed).toBe(1);const before=await d.save();expect(()=>p.styleId).toThrow('stale');await expect(span.replace('Updated')).rejects.toThrow('stale');expect(await d.save()).toEqual(before);
});
test('null removes only direct style without requiring a style registry',async()=>{
 const d=Document.create();d.addParagraph('x');const p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace(/(<w:p\b[^>]*>)/,'$1<w:pPr><w:pStyle w:val="Unknown"/><w:keepNext/></w:pPr>'));const q=await Document.open(p.toBytes());expect(q.paragraphs[0]!.styleId).toBe('Unknown');expect(q.paragraphs[0]!.setStyle(null)).toEqual({changed:1});const r=await OpcPackage.open(await q.save());expect(q.paragraphs[0]!.styleId).toBeUndefined();expect(r.text(r.mainPart())).toContain('<w:keepNext/>');expect(r.names()).not.toContain('word/styles.xml');
});

test('aliased namespaces and encoded style IDs retain text and unrelated properties',async()=>{
 const d=await styleDocument('assign'),p=await OpcPackage.open(await d.save());p.set('word/styles.xml',p.text('word/styles.xml').replace('w:styleId="Heading1"','w:styleId="A&amp;B"'));p.set(p.mainPart(),p.text(p.mainPart()).replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:'));const q=await Document.open(p.toBytes());expect(q.paragraphs[0]!.setStyle('A&B').changed).toBe(1);const saved=await OpcPackage.open(await q.save());expect(q.paragraphs[0]!.styleId).toBe('A&B');expect(saved.text(saved.mainPart())).toContain('w:val="A&amp;B"');expect(saved.text(saved.mainPart())).toContain('<q:keepNext/>');expect(saved.get('word/styles.xml')).toEqual(p.get('word/styles.xml'));
});
test('same direct style retains quote spelling and exact archive bytes',async()=>{
 const d=await styleDocument('same-style'),p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replace('<w:pStyle w:val="Heading1"/>',"<w:pStyle w:val = 'Heading&#49;' />"));const q=await Document.open(p.toBytes()),before=await q.save(),handle=q.paragraphs[0]!;expect(handle.setStyle('Heading1').changed).toBe(0);expect(await q.save()).toEqual(before);expect(handle.styleId).toBe('Heading1');
});
test('missing, self-closing properties and self-closing empty paragraphs accept first style child',async()=>{
 for(const shape of ['no-properties','empty-properties','empty-paragraph']){
  const d=await styleDocument(),p=await OpcPackage.open(await d.save());let xml=p.text(p.mainPart());xml=xml.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/,shape==='empty-properties'?'<w:pPr data="keep"/>':'');if(shape==='empty-paragraph')xml=xml.replace(/<w:p\b([^>]*)>[\s\S]*?<\/w:p>/,'<w:p$1/>');p.set(p.mainPart(),xml);const q=await Document.open(p.toBytes());expect(q.paragraphs[0]!.setStyle('Heading1').changed).toBe(1);const saved=await OpcPackage.open(await q.save());expect(q.paragraphs[0]!.styleId).toBe('Heading1');expect(q.paragraphs[0]!.text).toBe(shape==='empty-paragraph'?'':'Styled π text');if(shape==='empty-properties')expect(saved.text(saved.mainPart())).toContain('data="keep"');
 }
});
test('style-only update preserves UTF-16 encoding and BOM',async()=>{
 const d=await styleDocument(),p=await OpcPackage.open(await d.save()),xml=p.text(p.mainPart()).replace('encoding="UTF-8"','encoding="UTF-16"');const data=new Uint8Array(2+xml.length*2),view=new DataView(data.buffer);view.setUint16(0,0xfeff,true);for(let i=0;i<xml.length;i++)view.setUint16(2+2*i,xml.charCodeAt(i),true);p.set(p.mainPart(),data);const q=await Document.open(p.toBytes());q.paragraphs[0]!.setStyle('Heading1');const saved=await OpcPackage.open(await q.save());expect([...saved.get(saved.mainPart())!.slice(0,2)]).toEqual([255,254]);expect(saved.text(saved.mainPart())).toContain('encoding="UTF-16"');expect(q.paragraphs[0]!.styleId).toBe('Heading1');
});
test('serialization rollback leaves earlier changes and paragraph handles intact',async()=>{
 const d=await styleDocument();d.addParagraph('earlier');const before=await d.save(),handle=d.paragraphs[0]!,serialize=OpcPackage.prototype.toBytes;OpcPackage.prototype.toBytes=function(){throw new Error('injected style save failure');};try{expect(()=>handle.setStyle('Heading1')).toThrow('injected style save failure');}finally{OpcPackage.prototype.toBytes=serialize;}expect(await d.save()).toEqual(before);expect(handle.styleId).toBeUndefined();expect(handle.setStyle('Heading1').changed).toBe(1);
});
test('table handles survive style changes but stale cells and spans refuse',async()=>{
 const d=await styleDocument();const table=d.addTable(1,1);table.cell(0,0).text='cell';const cell=table.cell(0,0),p=d.paragraphs.find(p=>p.text==='cell')!,span=p.find('cell')[0]!;p.setStyle('Heading1');expect(()=>cell.text).toThrow('stale');await expect(span.replace('old')).rejects.toThrow('stale');expect(table.cell(0,0).text).toBe('cell');const q=await Document.open(await d.save());expect(q.tables[0]!.cell(0,0).text).toBe('cell');expect(q.paragraphs.find(p=>p.text==='cell')!.styleId).toBe('Heading1');
});
test('removal requests still refuse protection and invalid topology',async()=>{
 for(const kind of ['protected','duplicate-properties','duplicate-pstyle','misplaced-pstyle','wrong-namespace','property-revision']){const d=await styleDocument(kind),before=await d.save();expect(()=>d.paragraphs[0]!.setStyle(null)).toThrow();expect(await d.save()).toEqual(before);}
});
test('no registry, invalid registry and character styles refuse addition without touching the package',async()=>{
 const d=Document.create();d.addParagraph('x');const before=await d.save();expect(()=>d.paragraphs[0]!.setStyle('Heading1')).toThrow();expect(await d.save()).toEqual(before);
 for(const kind of ['character-style','duplicate-style','duplicate-relationship','external-styles','wrong-mime']){const q=await styleDocument(kind),snapshot=await q.save();expect(()=>q.addParagraph('new',{style:kind==='character-style'?'Emphasis':'Heading1'})).toThrow();expect(await q.save()).toEqual(snapshot);}
 const p=await OpcPackage.open(await(await styleDocument()).save());p.set('word/styles.xml','<wrong/>');const q=await Document.open(p.toBytes()),snapshot=await q.save();expect(()=>q.paragraphs[0]!.setStyle('Heading1')).toThrow();expect(await q.save()).toEqual(snapshot);
});
test('style getters inspect direct override only and do not select defaults',async()=>{
 const d=await styleDocument(),before=await d.save();expect(d.paragraphs[0]!.styleId).toBeUndefined();expect(await d.save()).toEqual(before);for(const value of ['', ' ',undefined,{},true]){expect(()=>d.paragraphs[0]!.setStyle(value as any)).toThrow();expect(await d.save()).toEqual(before);}
});

test('default namespace, direct run formatting and fresh spans survive style assignment',async()=>{
 const d=await styleDocument(),p=await OpcPackage.open(await d.save());p.set(p.mainPart(),p.text(p.mainPart()).replaceAll('w:','').replace('xmlns:w=','xmlns=').replace('val="','x:val="').replace('<document ','<document xmlns:x="'+W+'" '));const q=await Document.open(p.toBytes());q.paragraphs[0]!.setStyle('Heading1');const span=q.paragraphs[0]!.find('Styled')[0]!;await span.replace('Updated');expect(q.paragraphs[0]!.text).toBe('Updated π text');expect(q.paragraphs[0]!.styleId).toBe('Heading1');const saved=await OpcPackage.open(await q.save());expect(saved.text(saved.mainPart())).toContain('<b/>');
});
test('external or later enforcing settings refuse while explicit disabled settings allow selection',async()=>{
 for(const mode of ['external','mixed','disabled']){
  const d=await styleDocument(),p=await OpcPackage.open(await d.save());if(mode==='external')addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','https://example.invalid/settings.xml',{external:true});else{addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="0"/>${mode==='mixed'?'<w:documentProtection w:enforcement="1"/>':''}</w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');}
  const q=await Document.open(p.toBytes()),before=await q.save();if(mode==='disabled')expect(q.paragraphs[0]!.setStyle('Heading1').changed).toBe(1);else{expect(()=>q.paragraphs[0]!.setStyle('Heading1')).toThrow();expect(await q.save()).toEqual(before);}
 }
});
test('style mutations refuse lexical barriers, run revisions and unknown paragraph properties',async()=>{
 for(const [from,to]of [['<w:keepNext/>','<!--barrier--><w:keepNext/>'],['<w:keepNext/>','<w:unknown/>'],['<w:b/>','<w:b/><w:rPrChange/>'],['<w:t>Styled π text</w:t>','<w:instrText>Styled π text</w:instrText>']] as const){const d=await styleDocument(),p=await OpcPackage.open(await d.save());const original=p.text(p.mainPart());expect(original.includes(from)).toBe(true);p.set(p.mainPart(),original.replace(from,to));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.paragraphs[0]!.setStyle('Heading1')).toThrow();expect(await q.save()).toEqual(before);}
});
test('disk save and reopen retain chosen direct style and untouched definitions',async()=>{
 const {mkdtemp,rm}=await import('node:fs/promises'),{tmpdir}=await import('node:os'),root=await mkdtemp(join(tmpdir(),'bun-style-save-'));try{const d=await styleDocument(),styles=d.package.get('word/styles.xml')!;d.paragraphs[0]!.setStyle('Heading1');const path=join(root,'styled.docx');await d.save(path);const q=await Document.open(path);expect(q.paragraphs[0]!.styleId).toBe('Heading1');expect(q.package.get('word/styles.xml')).toEqual(styles);}finally{await rm(root,{recursive:true,force:true});}
});

test('same-style request still refuses document protection without touching handles',async()=>{
 const d=await styleDocument('same-style'),p=await OpcPackage.open(await d.save());addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const q=await Document.open(p.toBytes()),before=await q.save(),handle=q.paragraphs[0]!;expect(handle.styleId).toBe('Heading1');expect(()=>handle.setStyle('Heading1')).toThrow();expect(await q.save()).toEqual(before);expect(handle.styleId).toBe('Heading1');
});
