import assert from 'node:assert/strict';
import {mkdtemp,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
import {parseXml,type XmlElement} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export type ModelState={document:Document;body?:XmlElement;paragraphCount?:number;tableCount?:number;expected?:string[];before?:ReadonlyMap<string,Uint8Array>;reopened?:Document;destination?:string;disk?:Uint8Array};
const state=(c:Record<string,unknown>)=>c.state as ModelState;
export const scenarioIds=['@id-docx-go-new-empty-body','@id-docx-go-table-dimensions-getters','@id-docx-go-roundtrip-table-text'];
export const bindings:StepBinding[]=[
 {pattern:/^a new Word document$/,run:c=>{state(c).document=Document.create();}},
 {pattern:/^its body paragraphs and tables are enumerated$/,run:c=>{
  const s=state(c),bytes=s.document.package.get('word/document.xml');assert(bytes);const xml=parseXml(new TextDecoder('utf-8',{fatal:true}).decode(bytes));
  assert.equal(xml.root.namespaceURI,W);assert.equal(xml.root.localName,'document');const bodies=xml.root.children.filter(n=>n.namespaceURI===W&&n.localName==='body');assert.equal(bodies.length,1);
  s.body=bodies[0];s.paragraphCount=s.document.paragraphs.length;s.tableCount=s.document.tables.length;
 }},
 {pattern:/^the body is present with zero paragraphs and zero tables$/,run:c=>{
  const s=state(c);assert(s.body);assert.equal(s.paragraphCount,0);assert.equal(s.tableCount,0);assert.equal(s.body.children.filter(n=>n.namespaceURI===W&&(n.localName==='p'||n.localName==='tbl')).length,0);
 }},
 {pattern:/^a table with (\d+) rows and (\d+) columns is added$/,run:(c,rows,cols)=>{state(c).document.addTable(Number(rows),Number(cols));}},
 {pattern:/^RowCount equals (\d+) and ColumnCount equals (\d+) in memory$/,run:(c,rows,cols)=>{
  const tables=state(c).document.tables;assert.equal(tables.length,1);assert.equal(tables[0]!.rows,Number(rows));assert.equal(tables[0]!.columns,Number(cols));
 }},
 {pattern:/^a new Word table with three rows and three columns$/,run:c=>{const s=state(c);s.document=Document.create();s.document.addTable(3,3);}},
 {pattern:/^its cells contain Header1, Header2, Header3, A1, B1, C1, A2, B2 and C2 in row order$/,run:c=>{
  const s=state(c);s.expected=['Header1','Header2','Header3','A1','B1','C1','A2','B2','C2'];const table=s.document.tables[0];assert(table);assert.equal(table.rows,3);assert.equal(table.columns,3);
  for(let i=0;i<s.expected.length;i++)table.cell(Math.floor(i/3),i%3).text=s.expected[i]!;
  s.before=s.document.package.parts;
 }},
 {pattern:/^the document is saved and reopened$/,run:async c=>{
  const s=state(c),root=await mkdtemp(join(tmpdir(),'docx-model-'));try{
   s.destination=join(root,'table.docx');await s.document.save(s.destination);s.disk=Uint8Array.from(await Bun.file(s.destination).bytes());s.reopened=await Document.open(s.destination);
  }finally{await rm(root,{recursive:true,force:true});}
 }},
 {pattern:/^exactly one table is readable$/,run:c=>{const s=state(c);assert(s.reopened);assert.equal(s.reopened.tables.length,1);}},
 {pattern:/^all nine cell text getters equal their original row-order values$/,run:c=>{
  const s=state(c);assert(s.reopened&&s.expected&&s.before&&s.disk);assert.equal(s.expected.length,9);assert.equal(s.reopened.tables.length,1);const table=s.reopened.tables[0]!;assert.equal(table.rows,3);assert.equal(table.columns,3);
  for(let i=0;i<s.expected.length;i++)assert.equal(table.cell(Math.floor(i/3),i%3).text,s.expected[i]);
  const after=s.reopened.package.parts;assert.deepEqual([...after.keys()],[...s.before.keys()]);for(const[n,b]of s.before)if(n!=='word/document.xml')assert.deepEqual(after.get(n),b);
  assert.deepEqual(s.reopened.package.toBytes(),s.disk);
 }},
];
