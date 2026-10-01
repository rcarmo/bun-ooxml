import {test,expect} from 'bun:test';
import {predecessorPptxFeature} from '../helpers/pptx-shared-predecessor.ts';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/slide-order.ts';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
test('slide permutations execute21 saved outcomes and atomic refusals',async()=>{
 const path='workflows/pptx/slide-order.feature',f=parseFeature('references/fixtures-ooxml/'+path,(await predecessorPptxFeature(path,await Bun.file(join(fixturesRoot(),path)).text())).replace(/^@planned/m,'@implemented @bun')),rows=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(rows.length),steps:count(rows.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inv,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(21);
});

import {Presentation} from '../../src/pptx/index.ts';
import {OpcPackage} from '../../src/opc/index.ts';
import {orderDeck,P,R} from '../acceptance/slide-order.ts';
import {parseXml,elements} from '../../src/xml/index.ts';

test('reorder preserves captured slide anchors and table cells while indexes follow order',async()=>{
 const d=await orderDeck(),a=d.slides[0]!,b=d.slides[1]!;a.addTable(1,1,{x:0,y:0,width:1000,height:1000});a.tables[0]!.cell(0,0).text='before';const table=a.tables[0]!,cell=table.cell(0,0),anchor=a.inspectText('old')[0]!.anchor,version=d.currentSlideVersion(a.partName);d.reorderSlides([2,1,0]);expect(a.index).toBe(2);expect(b.index).toBe(1);expect(d.currentSlideVersion(a.partName)).toBe(version);expect(cell.text).toBe('before');cell.text='after';const fresh=a.inspectText('fresh')[0]!.anchor;a.replaceTextAt(fresh,'Alpha','Changed');expect(()=>a.replaceTextAt(anchor,'Alpha','stale')).toThrow('stale');a.addTextBox('same slide',{x:0,y:0,width:1000,height:1000});const q=await Presentation.open(d.package.toBytes());expect(q.slides[2]!.inspectText('saved').map(p=>p.text)).toEqual(['Changed','after','same slide']);
});
test('a text anchor captured before reorder can still edit its original slide',async()=>{
 const d=await orderDeck(),slide=d.slides[0]!,anchor=slide.inspectText('before')[0]!.anchor;d.reorderSlides([1,2,0]);slide.replaceTextAt(anchor,'Alpha','Updated');const q=await Presentation.open(d.package.toBytes());expect(q.slides.map(s=>s.inspectText('saved')[0]!.text)).toEqual(['Beta','Gamma','Updated']);
});
test('slide arrays are detached and repeated reorder plus append maintain indexes',async()=>{
 const d=await orderDeck(),a=d.slides[0]!,snapshot=d.slides;snapshot.reverse();snapshot.pop();expect(d.slides[0]).toBe(a);expect(d.slideCount).toBe(3);d.reorderSlides([2,0,1]);d.reorderSlides([1,2,0]);expect(d.slides[0]).toBe(a);expect(a.index).toBe(0);const last=d.addTextSlide('Delta');expect(last.index).toBe(3);d.reorderSlides([3,0,1,2]);expect(last.index).toBe(0);expect((await Presentation.open(d.package.toBytes())).slides.map(s=>s.inspectText('saved')[0]!.text)).toEqual(['Delta','Alpha','Beta','Gamma']);
});
test('serialization rollback preserves earlier edits, handle identities and order snapshot',async()=>{
 const d=await orderDeck(),a=d.slides[0]!;a.addTextBox('earlier',{x:0,y:0,width:1000,height:1000});const before=d.package.toBytes(),handles=d.slides,serialize=OpcPackage.prototype.toBytes;OpcPackage.prototype.toBytes=function(){throw Error('injected order serialization');};try{expect(()=>d.reorderSlides([2,1,0])).toThrow('injected order serialization');}finally{OpcPackage.prototype.toBytes=serialize;}expect(d.package.toBytes()).toEqual(before);expect(d.slides).toEqual(handles);expect(a.index).toBe(0);d.reorderSlides([2,1,0]);expect(a.index).toBe(2);expect(a.inspectText('after').at(-1)!.text).toBe('earlier');
});
test('UTF16 LE/BE presentation metadata preserves BOM and all other payloads',async()=>{
 for(const le of [true,false]){const d=await orderDeck(),main=d.package.mainPart(),xml=d.package.text(main).replace('encoding="UTF-8"','encoding="UTF-16"'),bytes=new Uint8Array(2+xml.length*2),view=new DataView(bytes.buffer);view.setUint16(0,0xfeff,le);for(let i=0;i<xml.length;i++)view.setUint16(2+2*i,xml.charCodeAt(i),le);d.package.set(main,bytes);const q=await Presentation.open(d.package.toBytes()),old=new Map(q.package.names().map(n=>[n,q.package.get(n)!]));q.reorderSlides([2,1,0]);const saved=await Presentation.open(q.package.toBytes());expect([...saved.package.get(main)!.slice(0,2)]).toEqual(le?[255,254]:[254,255]);expect(saved.package.text(main)).toContain('encoding="UTF-16"');for(const [n,b]of old)if(n!==main)expect(saved.package.get(n)).toEqual(b);expect(saved.slides[0]!.inspectText('read')[0]!.text).toBe('Gamma');}
});
test('slide entries retain exact quote spelling, local aliases and whitespace slots',async()=>{
 const d=await orderDeck(),main=d.package.mainPart(),xml=d.package.text(main),nodes=elements(parseXml(xml),'sldId',P);let next=xml;for(const n of [...nodes].reverse()){const raw=xml.slice(n.start,n.end).replace('r:id=',`xmlns:z="${R}" z:id=`).replace(/id="(\d+)"/,"id = '$1'");next=next.slice(0,n.start)+'\n  '+raw+next.slice(n.end);}d.package.set(main,next);const q=await Presentation.open(d.package.toBytes()),list=elements(parseXml(next),'sldIdLst',P)[0]!,entries=list.children.map(n=>next.slice(n.start,n.end));q.reorderSlides([2,0,1]);const actual=q.package.text(main),list2=elements(parseXml(actual),'sldIdLst',P)[0]!;expect(list2.children.map(n=>actual.slice(n.start,n.end))).toEqual([entries[2]!,entries[0]!,entries[1]!]);expect(actual.slice(list2.openEnd,list2.closeStart)).toBe('\n  '+entries[2]+'\n  '+entries[0]+'\n  '+entries[1]);
});
test('permutations reject sparse/getter/nonarray/extended data without invoking getters',async()=>{
 const d=await orderDeck(),before=d.package.toBytes();let calls=0;const accessor=[0,1,2];Object.defineProperty(accessor,'1',{get(){calls++;return 1;}});const extra=[0,1,2];(extra as any).extra=1;
 for(const order of [null,{},[0,,2],accessor,extra,[0,1,'2'],[0,1,NaN],[0,1,Infinity],[-1,1,2],Object.assign([0,1,2],{[Symbol('extra')]:1})]){expect(()=>d.reorderSlides(order as any)).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(calls).toBe(0);
});
test('same-order requests still validate protection, IDs and stale metadata',async()=>{
 for(const kind of ['protected','duplicate-id','wrong-mime','extension-metadata','custom-show']){const d=await orderDeck(kind),before=d.package.toBytes();expect(()=>d.reorderSlides([0,1,2])).toThrow();expect(d.package.toBytes()).toEqual(before);}
 const d=await orderDeck(),root='_rels/.rels';d.package.set(root,d.package.text(root)+' ');const before=d.package.toBytes();expect(()=>d.reorderSlides([0,1,2])).toThrow('outside');expect(()=>d.addTextSlide('unsafe')).toThrow('outside');expect(d.package.toBytes()).toEqual(before);
});
test('invalid lexical slide IDs refuse before archive changes',async()=>{
 for(const id of ['255','2147483648','256junk','-256']){const d=await orderDeck(),main=d.package.mainPart(),xml=d.package.text(main),node=elements(parseXml(xml),'sldId',P)[0]!;d.package.set(main,xml.slice(0,node.start)+xml.slice(node.start,node.end).replace(/id="\d+"/,`id="${id}"`)+xml.slice(node.end));const q=await Presentation.open(d.package.toBytes()),before=q.package.toBytes();expect(()=>q.reorderSlides([2,1,0])).toThrow();expect(q.package.toBytes()).toEqual(before);}
});
test('disk saved slide order and linked notes follow logical handles',async()=>{
 const {mkdtemp,rm}=await import('node:fs/promises'),{tmpdir}=await import('node:os'),{join}=await import('node:path'),root=await mkdtemp(join(tmpdir(),'bun-slide-order-'));try{const d=await orderDeck('notes');d.reorderSlides([1,2,0]);const path=join(root,'ordered.pptx');await d.save(path);const q=await Presentation.open(path);expect(q.slides.map(s=>s.inspectText('saved')[0]!.text)).toEqual(['Beta','Gamma','Alpha']);expect(q.slides.map(s=>s.readNotesText())).toEqual(['Notes 1','Notes 2','Notes 0']);}finally{await rm(root,{recursive:true,force:true});}
});

test('noncanonical numeric own keys and accessors refuse instead of hiding extra order data',async()=>{
 const d=await orderDeck(),before=d.package.toBytes();let calls=0;for(const key of ['00','01','0002'])for(const accessor of [false,true]){const order=[2,1,0];Object.defineProperty(order,key,accessor?{get(){calls++;return 1;}}:{value:1});expect(()=>d.reorderSlides(order)).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(calls).toBe(0);
});

test('public three-argument Slide constructor retains identity across reorders', async () => {
 const {Slide}=await import('../../src/pptx/index.ts');
 const d=await orderDeck(),part=d.slides[0]!.partName;
 // The third constructor argument was public before reordering existed.
 const alias=new Slide(d,part,0);
 expect(alias.index).toBe(0);
 const anchor=alias.inspectText('constructor')[0]!.anchor;
 d.reorderSlides([2,1,0]);expect(alias.index).toBe(2);
 alias.replaceTextAt(anchor,'Alpha','Alias');
 alias.addTextBox('retained',{x:0,y:0,width:1000,height:1000});
 const q=await Presentation.open(d.package.toBytes());
 expect(q.slides[2]!.inspectText('saved').map(p=>p.text)).toEqual(['Alias','retained']);
 expect(()=>new Slide(d,'ppt/slides/absent.xml').index).toThrow('not in this presentation');
});
