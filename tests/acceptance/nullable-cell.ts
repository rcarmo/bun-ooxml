import assert from 'node:assert/strict';
import {Document,TableCell} from '../../src/docx/index.ts';
import type {AcceptanceStep,StepBinding} from '../../scripts/gherkin.ts';
export const scenarioIds=['@id-docx-go-table-cell-access'];
const doc=(c:Record<string,unknown>)=>(c.state as {document:Document}).document;
export const bindings:StepBinding[]=[
 {pattern:/^a new Word table has three rows and three columns with texts by row A1,B1,C1 then A2,B2,C2 then A3,B3,C3$/,run:c=>{
  const d=Document.create(),t=d.addTable(3,3);for(let r=0;r<3;r++)for(let col=0;col<3;col++)t.cell(r,col).text=String.fromCharCode(65+col)+(r+1);
  c.state={document:d};c.lookupBytes=d.package.get('word/document.xml')!.slice();c.lookupTexts=Array.from({length:3},(_,r)=>Array.from({length:3},(_,col)=>t.cell(r,col).text)).flat();
 }},
 {pattern:/^cells are looked up at these zero-based coordinates$/,run:c=>{
  const rows=(c.step as AcceptanceStep).argument?.dataTable;assert(rows);assert.deepEqual(rows[0],['row','column','present','text']);
  assert.equal(rows.length,15);c.lookupResults=rows.slice(1).map(row=>{assert.equal(row.length,4);const [r,col,present,text]=row;const cell=doc(c).tables[0]!.tryCell(Number(r),Number(col));return{present:cell!==null,text:cell?.text??'',expectedPresent:present==='true',expectedText:text};});
 }},
 {pattern:/^each lookup returns the listed presence and exact text without an exception$/,run:c=>{
  const rows=c.lookupResults as Array<{present:boolean;text:string;expectedPresent:boolean;expectedText:string}>;assert.equal(rows.length,14);
  assert.equal(rows.filter(r=>r.expectedPresent).length,9);for(const r of rows){assert.equal(r.present,r.expectedPresent);assert.equal(r.text,r.expectedText);}
 }},
 {pattern:/^the table still has three rows and three columns with its original texts and unchanged document XML$/,run:c=>{
  const d=doc(c),t=d.tables[0]!;assert.equal(d.tables.length,1);assert.equal(t.rows,3);assert.equal(t.columns,3);
  assert.deepEqual(Array.from({length:3},(_,r)=>Array.from({length:3},(_,col)=>t.cell(r,col).text)).flat(),c.lookupTexts);
  assert.deepEqual(d.package.get('word/document.xml'),c.lookupBytes);
 }},

 {pattern:/^its Cell getter is called for all nine coordinates from zero through two$/,run:c=>{const tables=doc(c).tables;assert.equal(tables.length,1);assert.equal(tables[0]!.rows,3);assert.equal(tables[0]!.columns,3);c.nullableCells=Array.from({length:3},(_,r)=>Array.from({length:3},(_,col)=>tables[0]!.tryCell(r,col))).flat();}},
 {pattern:/^each of those nine calls returns a nonnil cell$/,run:c=>{const cells=c.nullableCells as Array<TableCell|null>;assert.equal(cells.length,9);for(const cell of cells)assert(cell instanceof TableCell);}},
 {pattern:/^calls for row or column negative one or three at the tested boundary coordinates return nil$/,run:c=>{const t=doc(c).tables[0]!;for(const[r,col]of [[-1,0],[0,-1],[3,0],[0,3]])assert.equal(t.tryCell(r!,col!),null);}},
];
