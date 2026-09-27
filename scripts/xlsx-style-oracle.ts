/** Development-only independent-reader protocol and authored sample. Never imported by src. */
import assert from 'node:assert/strict';
import {join,basename} from 'node:path';
import {Workbook,OpcPackage,patchOffice} from '../src/index.ts';
import {elements,parseXml} from '../src/xml/index.ts';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const sha=(bytes:Uint8Array)=>new Bun.CryptoHasher('sha256').update(bytes).digest('hex');
export const styleCells=[
 {sheet:'Sheet1',address:'A1',value:'Independent\nwrapped value',styleIndex:1,explicitStyleIndex:true,wrapText:true,fontId:0,fillId:0,borderId:0,numberFormatId:0,baseStyleIndex:0},
 {sheet:'Sheet1',address:'B1',value:'42',styleIndex:0,explicitStyleIndex:true,wrapText:false,fontId:0,fillId:0,borderId:0,numberFormatId:0,baseStyleIndex:0},
 {sheet:'Audit',address:'A1',value:'Unchanged',styleIndex:0,explicitStyleIndex:false,wrapText:false,fontId:0,fillId:0,borderId:0,numberFormatId:0,baseStyleIndex:0},
];
export async function authorStyleSample(directory:string){
 const source=join(directory,'style-source.xlsx'),output=join(directory,'style-wrapped.xlsx');
 const book=Workbook.create();book.worksheet('Sheet1').setCellValue('A1','Style marker');book.worksheet('Sheet1').setCellValue('B1',42);book.worksheet('Sheet1').setCellStyle('B1',0);book.addWorksheet('Audit').setCellValue('A1','Unchanged');
 await book.save(source);const before=await Bun.file(source).bytes(),old=await OpcPackage.open(before);
 const count=(p:OpcPackage)=>elements(parseXml(p.text('xl/styles.xml')),'cellXfs',S)[0]!.children.length;
 assert.equal(count(old),1);
 const receipt=await patchOffice({source,output,mode:'strict',multilineWrap:true,changes:[{target:'Sheet1!A1',value:'Independent\nwrapped value'}]});
 assert.equal(receipt.status,'committed');assert.equal(receipt.results[0]!.matched,1);assert.equal(receipt.committedChanges,1);
 const bytes=await Bun.file(output).bytes(),after=await OpcPackage.open(bytes);assert.equal(count(after),2);assert.deepEqual(await Bun.file(source).bytes(),before);assert.deepEqual(after.names(),old.names());
 for(const name of old.names())if(!['xl/styles.xml','xl/worksheets/sheet1.xml'].includes(name))assert.deepEqual(after.get(name),old.get(name));
 const reopened=await Workbook.open(output);assert.equal(reopened.worksheet('Sheet1').getCell('A1')?.value,'Independent\nwrapped value');assert.equal(reopened.worksheet('Sheet1').getCell('A1')?.styleId,'1');
 return {source,output,sourceSha256:sha(before),sha256:sha(bytes),receipt};
}
export const styleFaults=[
 {name:'cell-index',part:'xl/worksheets/sheet1.xml',from:'s="1"',to:'s="999"',error:'cell format index 999'},
 ...(['font','fill','border'] as const).map(kind=>({name:kind+'-index',part:'xl/styles.xml',from:kind+'Id="0"',to:kind+'Id="999"',error:kind+' index 999'})),
 {name:'base-index',part:'xl/styles.xml',from:'xfId="0"',to:'xfId="999"',error:'base style index 999'},
 {name:'number-format',part:'xl/styles.xml',from:'numFmtId="0"',to:'numFmtId="999"',error:'custom number format 999'},
];
export async function corruptStyleSample(bytes:Uint8Array,fault:typeof styleFaults[number]){
 const pkg=await OpcPackage.open(bytes),old=pkg.text(fault.part),changed=old.replaceAll(fault.from,fault.to);assert.notEqual(changed,old,'Control must alter its intended field');pkg.set(fault.part,changed);return pkg.toBytes();
}
export function verifyStyleRefusal(report:any,path:string,hash:string,error:string){
 assert.equal(report.schemaVersion,1);assert.equal(report.validator,'DocumentFormat.OpenXml');assert.match(report.version,/^3\.5\.1(?:\+|$)/);assert.equal(report.status,'failed');assert.equal(report.profile,'Office2019');assert.equal(report.results?.length,1);
 const r=report.results[0];assert.equal(r.file,basename(path));assert.equal(r.sha256,hash);assert.equal(typeof r.exception,'string');assert.ok(r.exception.startsWith('styles: '));assert.ok(r.exception.includes(error));assert.equal(r.spreadsheetStyles,undefined);
}
export function verifyStyleReadback(report:any,path:string,hash:string){
 assert.equal(report.schemaVersion,1);assert.equal(report.validator,'DocumentFormat.OpenXml');assert.match(report.version,/^3\.5\.1(?:\+|$)/);assert.equal(report.status,'passed');assert.equal(report.profile,'Office2019');assert.equal(report.results?.length,1);
 const r=report.results[0];assert.equal(r.file,basename(path));assert.equal(r.sha256,hash);assert.deepEqual(r.errors,[]);assert.equal(r.truncated,false);assert.equal(r.exception,undefined);
 assert.deepEqual(r.spreadsheetStyles,{cellFormatCount:2,cells:styleCells});return r.spreadsheetStyles;
}
