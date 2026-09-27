import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
import type {ModelState} from './docx-model.ts';
type State=ModelState&{headerBefore?:boolean;headerAfter?:boolean};
const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-docx-go-table-header-getter'];
export const bindings:StepBinding[]=[
 {pattern:/^the first row of a new Word three-by-two table$/,run:c=>{const s=state(c);s.document=Document.create();s.document.addTable(3,2);}},
 {pattern:/^IsHeader is read, SetHeader true is applied and IsHeader is read again$/,run:c=>{const s=state(c),table=s.document.tables[0]!;s.headerBefore=table.isRowHeader(0);assert.equal(table.setRowHeader(0,true).changed,1);s.headerAfter=table.isRowHeader(0);}},
 {pattern:/^the first result is (true|false) and the second is (true|false)$/,run:(c,before,after)=>{const s=state(c);assert.equal(s.headerBefore,before==='true');assert.equal(s.headerAfter,after==='true');}},
];
