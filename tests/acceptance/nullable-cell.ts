import assert from 'node:assert/strict';
import {Document,TableCell} from '../../src/docx/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const scenarioIds=['@id-docx-go-table-cell-access'];
const doc=(c:Record<string,unknown>)=>(c.state as {document:Document}).document;
export const bindings:StepBinding[]=[
 {pattern:/^its Cell getter is called for all nine coordinates from zero through two$/,run:c=>{const tables=doc(c).tables;assert.equal(tables.length,1);assert.equal(tables[0]!.rows,3);assert.equal(tables[0]!.columns,3);c.nullableCells=Array.from({length:3},(_,r)=>Array.from({length:3},(_,col)=>tables[0]!.tryCell(r,col))).flat();}},
 {pattern:/^each of those nine calls returns a nonnil cell$/,run:c=>{const cells=c.nullableCells as Array<TableCell|null>;assert.equal(cells.length,9);for(const cell of cells)assert(cell instanceof TableCell);}},
 {pattern:/^calls for row or column negative one or three at the tested boundary coordinates return nil$/,run:c=>{const t=doc(c).tables[0]!;for(const[r,col]of [[-1,0],[0,-1],[3,0],[0,3]])assert.equal(t.tryCell(r!,col!),null);}},
];
