import {expect} from 'bun:test';
import {Document,type Paragraph} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export async function formattingDocument(kind:string){
 const d=Document.create();d.addParagraph('Alpha');const pkg=await OpcPackage.open(d.package.toBytes()),main=pkg.mainPart();let xml=pkg.text(main).replace(`<w:p xmlns:w="${W}">`,'<w:p>'),selected='',styles='',defaults='',direct='';
 const style=(id:string,body:string,attrs='')=>`<w:style w:type="paragraph" w:styleId="${id}" ${attrs}>${body}</w:style>`;
 if(kind==='defaults')defaults='<w:docDefaults><w:rPrDefault><w:rPr><w:b/><w:i w:val="false"/></w:rPr></w:rPrDefault></w:docDefaults>';
 if(kind==='default-style')styles=style('Normal','<w:rPr><w:b/></w:rPr>','w:default="1"');
 if(!['implicit','defaults','default-style','table-context'].includes(kind)){selected='Child';styles=style('Base','<w:rPr><w:b/></w:rPr>')+style('Child','<w:basedOn w:val="Base"/><w:rPr><w:i/></w:rPr>');}
 if(kind==='toggle-chain')styles=style('Base','<w:rPr><w:b/></w:rPr>')+style('Child','<w:basedOn w:val="Base"/><w:rPr><w:b/><w:i/></w:rPr>');
 if(kind==='style-false')styles=style('Base','<w:rPr><w:b/></w:rPr>')+style('Child','<w:basedOn w:val="Base"/><w:rPr><w:b w:val="0"/><w:i/></w:rPr>');
 if(kind==='direct-off')direct='<w:b w:val="0"/><w:i w:val="false"/>';
 if(kind==='missing-style')styles='';
 if(kind==='cycle')styles=style('Base','<w:basedOn w:val="Child"/>')+style('Child','<w:basedOn w:val="Base"/>');
 if(kind==='duplicate-id')styles+=style('Child','');
 if(kind==='duplicate-default')styles+=style('N1','','w:default="1"')+style('N2','','w:default="true"');
 if(kind==='wrong-type')styles=styles.replace('<w:style w:type="paragraph" w:styleId="Child"','<w:style w:type="character" w:styleId="Child"');
 if(kind==='malformed-flag')styles=styles.replace('<w:b/>','<w:b w:val="maybe"/>');
 if(kind==='duplicate-flag')styles=styles.replace('<w:b/>','<w:b/><w:b/>');
 if(kind==='character-style')direct='<w:rStyle w:val="Emphasis"/>';
 if(kind==='complex-script')direct='<w:cs/>';
 if(kind==='revision')direct='<w:rPrChange w:id="0"/>';
 if(selected)xml=xml.replace('<w:p>','<w:p><w:pPr><w:pStyle w:val="'+selected+'"/>'+(kind==='numbering'?'<w:numPr/>':'')+'</w:pPr>');
 if(direct)xml=xml.replace('<w:r>','<w:r><w:rPr>'+direct+'</w:rPr>');
 if(kind==='multiple-runs')xml=xml.replace('</w:p>','<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>Beta</w:t></w:r></w:p>');
 if(kind==='table-context')xml=xml.replace('<w:p>','<w:tbl><w:tr><w:tc><w:p>').replace('</w:p>','</w:p></w:tc></w:tr></w:tbl>');
 let registry=`<w:styles xmlns:w="${W}">${defaults}${styles}</w:styles>`;
 if(kind==='aliased'){xml=xml.replaceAll('xmlns:w=','xmlns:q=').replace(/(<\/?|\s)w:/g,'$1q:');registry=registry.replaceAll('xmlns:w=','xmlns:q=').replace(/(<\/?|\s)w:/g,'$1q:');}
 pkg.set(main,xml);
 if(kind!=='implicit'&&kind!=='table-context'){
  addPart(pkg,'word/styles.xml',registry,kind==='wrong-mime'?'application/xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
  addRelationship(pkg,main,R+'/styles',kind==='external-styles'?'https://example.invalid/styles.xml':'styles.xml',{external:kind==='external-styles'});
 }
 if(kind==='styles-effects')addRelationship(pkg,main,'http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects','styles.xml');
 addPart(pkg,'custom/formatting-sentinel.bin',new Uint8Array([7,9,11]),'application/octet-stream');
 return Document.open(pkg.toBytes());
}
export const bindings:StepBinding[]=[
 {pattern:/^a document prepared for effective formatting (\S+)$/,run:async(c,k)=>{const d=await formattingDocument(k!);c.doc=d;c.kind=k;c.handle=d.paragraphs[0]!;c.before=d.package.toBytes();c.version=d.currentVersion();}},
 {pattern:/^an unsafe effective-formatting document (\S+)$/,run:async(c,k)=>{const d=await formattingDocument(k!);c.doc=d;c.kind=k;c.handle=d.paragraphs[0]!;if(k==='stale-paragraph')d.addParagraph('later');c.before=d.package.toBytes();}},
 {pattern:/^its plain paragraph run formatting is inspected$/,run:c=>{try{c.result=(c.handle as any).effectiveRunFormatting();}catch(e){c.error=e;}}},
 {pattern:/^effective flags and provenance match (\S+) after reopening$/,run:async(c,k)=>{
  expect(c.error).toBeUndefined();const rows=c.result as any[];
  const expected=k==='implicit'?[[false,false]]:k==='defaults'||k==='default-style'?[[true,false]]:k==='toggle-chain'?[[false,true]]:k==='direct-off'?[[false,false]]:k==='multiple-runs'?[[true,true],[false,true]]:[[true,true]];
  expect(rows.map(r=>[r.bold.value,r.italic.value])).toEqual(expected);expect(rows.map(r=>r.text)).toEqual(k==='multiple-runs'?['Alpha','Beta']:['Alpha']);
  const ids=k==='implicit'||k==='defaults'?[]:k==='default-style'?['Normal']:['Base','Child'];expect(rows[0].paragraphStyleChain).toEqual(ids);
  expect(rows[0].bold.contributions[0]).toEqual({source:'implicit',operation:'set',value:false,result:false});
  if(k==='toggle-chain')expect(rows[0].bold.contributions.slice(1)).toEqual([{source:'paragraph-style',styleId:'Base',partName:'word/styles.xml',operation:'toggle',value:true,result:true},{source:'paragraph-style',styleId:'Child',partName:'word/styles.xml',operation:'toggle',value:true,result:false}]);
  if(k==='direct-off')expect(rows[0].bold.contributions.at(-1)).toEqual({source:'direct-run',partName:'word/document.xml',operation:'set',value:false,result:false});
  const d=c.doc as Document,q=await Document.open(d.package.toBytes());expect((q.paragraphs[0] as any).effectiveRunFormatting()).toEqual(rows);
 }},
 {pattern:/^formatting inspection preserves all bytes and existing handles$/,run:c=>{const d=c.doc as Document;expect(d.package.toBytes()).toEqual(c.before as Uint8Array);expect(d.paragraphs[0]).toBe(c.handle as Paragraph);expect(d.currentVersion()).toBe(c.version as number);}},
 {pattern:/^effective formatting inspection refuses without changing package bytes$/,run:c=>{expect(c.error).toBeDefined();expect((c.error as any).code).toMatch(/^docx-/);expect((c.doc as Document).package.toBytes()).toEqual(c.before as Uint8Array);}},
];
