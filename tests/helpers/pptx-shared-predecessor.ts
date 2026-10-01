import assert from 'node:assert/strict';
import {join} from 'node:path';
import type {AcceptanceInventory} from '../../scripts/gherkin.ts';
import pin from '../../docs/behaviors/pptx-manipulation-candidate.json';
import formatting from '../../docs/behaviors/pptx-formatting-candidate.json';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
/** Historical profile tests exclude only the reviewed additive PPTX cases. */
export function historicalSharedCases(inv:AcceptanceInventory){const cases=inv.coverage!.shared.cases,added=inv.features.flatMap(f=>f.scenarios).filter(s=>[...pin.selectedScenarioIds,...formatting.selectedScenarioIds].includes(s.scenarioId));if(!added.length)return cases;assert([20,40].includes(added.length));assert(added.every(s=>s.lifecycle==='implemented'));return {...cases,implemented:cases.implemented-added.length,total:cases.total-added.length};}
/** Exact predecessor recipe for isolated old bindings and assertion ledgers. */
export async function predecessorPptxFeature(path:string,text:string){if(!process.env.OOXML_REFERENCE_PIN)return text;const commit=(await Bun.file(process.env.OOXML_REFERENCE_PIN).json()).commit;if(commit!==pin.commit&&commit!==formatting.commit)return text;if(commit===formatting.commit){const l=await Bun.file(join(fixturesRoot(),'ledgers/pptx-formatting.json')).json(),f=l.files.find((f:any)=>f.path===path);if(f){assert.equal(new Bun.CryptoHasher('sha256').update(text).digest('hex'),f.afterSha256);text=f.beforeText;}}const ledger=await Bun.file(join(fixturesRoot(),'ledgers/pptx-manipulation.json')).json(),f=ledger.files.find((f:any)=>f.path===path);if(!f)return text;assert.equal(new Bun.CryptoHasher('sha256').update(text).digest('hex'),f.afterSha256);return f.beforeText;}
