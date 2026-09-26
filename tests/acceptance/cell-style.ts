import {expect} from 'bun:test';
import {Workbook} from '../../src/xlsx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',STYLE_REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',STYLE_TYPE='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
export const part='xl/worksheets/sheet1.xml';
export async function styledWorkbook(kind='assign'){
 const w=Workbook.create();w.worksheet('Sheet1').setCellValue('A1','value π');w.addWorksheet('Other').setCellValue('B1',7);const p=w.package;
 let styles=p.text('xl/styles.xml').replace('<cellXfs count="1">','<cellXfs count="2">').replace('</cellXfs>','<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>');
 if(kind==='wrong-root')styles='<wrong/>';
 if(kind==='duplicate-cellxfs')styles=styles.replace('</styleSheet>','<cellXfs count="0"/></styleSheet>');
 if(kind==='incorrect-count')styles=styles.replace('cellXfs count="2"','cellXfs count="3"');
 if(kind==='invalid-font')styles=styles.replace('applyAlignment="1"','fontId="99"').replace('fontId="0" fillId="0" borderId="0" xfId="0" fontId="99"','fontId="99" fillId="0" borderId="0" xfId="0"');
 if(kind==='invalid-base')styles=styles.replace('xfId="0" applyAlignment="1"','xfId="9" applyAlignment="1"');
 if(kind==='missing-number-format')styles=styles.replace('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment','<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment');
 p.set('xl/styles.xml',styles);
 let xml=p.text(part);const style=['replace','remove'].includes(kind)?'0':kind==='no-op'?'1':kind==='malformed-old-index'?'oops':undefined;
 if(style!==undefined)xml=xml.replace('r="A1"',`r="A1" s="${style}"`);
 if(kind==='formula')xml=xml.replace(/<c\b[^>]*>[\s\S]*?<\/c>/,'<c r="A1"><f>Other!B1+1</f><v>8</v></c>');
 if(kind==='blank')xml=xml.replace(/<c\b[^>]*>[\s\S]*?<\/c>/,'<c r="A1"/>');
 if(kind==='aliased')xml=xml.replace('xmlns="'+S+'"','xmlns:q="'+S+'"').replace(/<(\/?)(worksheet|sheetData|row|c|is|t|dimension)(?=[\s/>])/g,'<$1q:$2');
 if(kind==='protected-sheet')xml=xml.replace('</worksheet>','<sheetProtection sheet="1"/></worksheet>');p.set(part,xml);
 if(kind==='protected-workbook'){const main=p.mainPart();p.set(main,p.text(main).replace('</workbook>','<workbookProtection lockStructure="1"/></workbook>'));}
 if(kind==='duplicate-relationship'){addPart(p,'xl/styles2.xml',styles,STYLE_TYPE);addRelationship(p,p.mainPart(),STYLE_REL,'styles2.xml');}
 if(kind==='external-styles'){const rels='xl/_rels/workbook.xml.rels';p.set(rels,p.text(rels).replace('Target="styles.xml"','Target="https://example.invalid/styles.xml" TargetMode="External"'));}
 if(kind==='worksheet-alias'){const rel='xl/_rels/workbook.xml.rels';p.set(rel,p.text(rel).replace('worksheets/sheet2.xml','worksheets/sheet1.xml'));}
 if(kind==='wrong-mime')p.set('[Content_Types].xml',p.text('[Content_Types].xml').replace(STYLE_TYPE,'application/xml'));
 return Workbook.open(p.toBytes());
}
export const bindings:StepBinding[]=[
 {pattern:/^a native workbook prepared for cell-style (\S+)$/,run:async(c,k)=>{c.kind=k;c.doc=await styledWorkbook(k);c.before=(c.doc as Workbook).toBytes();}},
 {pattern:/^the existing cell style is selected for (\S+)$/,run:(c,k)=>{c.receipt=(c.doc as Workbook).worksheet('Sheet1').setCellStyle('A1',k==='remove'||k==='absent-removal'?null:k==='explicit-zero'?0:1);}},
 {pattern:/^reopened cell style and change receipt match (\S+)$/,run:async(c,k)=>{const w=await Workbook.open((c.doc as Workbook).toBytes());expect(w.worksheet('Sheet1').getCell('A1')!.styleId).toBe(k==='remove'||k==='absent-removal'?undefined:k==='explicit-zero'?'0':'1');expect(c.receipt).toEqual({changed:k==='no-op'||k==='absent-removal'?0:1});}},
 {pattern:/^values formulas caches and unrelated package parts retain exact custody$/,run:async c=>{const w=c.doc as Workbook,before=await OpcPackage.open(c.before as Uint8Array);expect(w.package.names()).toEqual(before.names());for(const n of before.names())if(n!==part)expect(w.package.get(n)).toEqual(before.get(n));const strip=(xml:string)=>{const node=elements(parseXml(xml),'c',S)[0]!;return xml.slice(0,node.start)+xml.slice(node.start,node.openEnd).replace(/\s+s="[^"]*"/,'')+xml.slice(node.openEnd);};expect(strip(w.package.text(part))).toBe(strip(before.text(part)));if(c.kind==='no-op'||c.kind==='absent-removal')expect(w.toBytes()).toEqual(c.before as Uint8Array);}},
 {pattern:/^an unsafe cell-style selection input (\S+)$/,run:async(c,k)=>{const w=await styledWorkbook(k);c.kind=k;c.doc=w;c.cell=w.worksheet('Sheet1').getCell('A1');if(k==='stale-sheet')w.package.set(part,w.package.text(part).replace('value π','external'));if(k==='stale-relationship'){const rel='xl/_rels/workbook.xml.rels';w.package.set(rel,w.package.text(rel)+' ');}c.before=w.toBytes();}},
 {pattern:/^its existing cell-style selection is attempted$/,run:c=>{try{c.receipt=(c.doc as Workbook).worksheet('Sheet1').setCellStyle(c.kind==='missing-cell'?'B2':'A1',c.kind==='invalid-index'?-1:c.kind==='missing-index'?8:1);}catch(e){c.error=e;}}},
 {pattern:/^cell-style selection refuses without changing package bytes or cached cell state$/,run:c=>{expect(c.error).toBeDefined();expect((c.error as any).code).toMatch(/^xlsx-/);const w=c.doc as Workbook;expect(w.toBytes()).toEqual(Uint8Array.from(c.before as Uint8Array));expect(w.worksheet('Sheet1').getCell('A1')).toEqual(c.cell as any);}},
];
