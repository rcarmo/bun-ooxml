import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {selectSharedScenarios,type AcceptanceFeature} from '../../scripts/gherkin.ts';
/** Resolve current canonical locations by identity; never reads a retired feature. */
export async function sharedScenarios(ids:readonly string[],change:(text:string)=>string=text=>text):Promise<AcceptanceFeature[]>{
 if(!ids.length||new Set(ids).size!==ids.length)throw Error('Expected nonempty unique scenario IDs');
 const ledger=await Bun.file(join(fixturesRoot(),'ledgers/workflows.json')).json();
 const owners=new Map<string,string[]>();
 for(const id of ids){const rows=ledger.workflows.filter((w:any)=>w.id===id);if(rows.length!==1)throw Error('Expected one shared scenario owner: '+id);const path=rows[0].feature;owners.set(path,[...(owners.get(path)??[]),id]);}
 return Promise.all([...owners].map(async([path,selected])=>{
  const feature=selectSharedScenarios(path,change(await Bun.file(join(fixturesRoot(),path)).text()),selected);
  return {...feature,scenarios:feature.scenarios.filter(s=>selected.includes(s.scenarioId))};
 }));
}
