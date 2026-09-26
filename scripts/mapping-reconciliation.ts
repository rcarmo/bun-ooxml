/** @script Reconcile historical Bun staging IDs against the native declaration inventory.
 * @usage bun scripts/mapping-reconciliation.ts [--check]
 * @description Identity comparison only; no source-equivalence or execution credit.
 */
import {join,resolve} from 'node:path';
import {inventoryNativeTests,type TestCaseRecord} from './test-inventory.ts';
const hash=(s:string)=>new Bun.CryptoHasher('sha256').update(s).digest('hex');
export function reconcile(cases:TestCaseRecord[],files:{path:string;source:string}[]){
 const current=new Map(cases.map(c=>[c.id,c])),seen=new Set<string>();
 const mappings=files.flatMap(file=>{
  const data=JSON.parse(file.source);if(data.reviewState!=='candidate-needs-parent-review'||!Array.isArray(data.mappings))throw Error('Unexpected staging mapping schema: '+file.path);
  return data.mappings.map((row:any)=>{
   if(typeof row.testId!=='string'||!row.testId.startsWith('bun:')||seen.has(row.testId))throw Error('Invalid or duplicate staging identity');seen.add(row.testId);
   const match=current.get(row.testId);
   return {testId:row.testId,stagingPath:file.path,state:match?'identity-present-needs-review':'identity-missing',sourceEquivalence:'unverified-staging-has-no-source-pin',currentSourceSha256:match?.sourceSha256??null,currentReviewReasons:match?.reviewReasons??[],executionCredit:false};
  });
 }).sort((a,b)=>a.testId.localeCompare(b.testId));
 const withoutStaging=cases.filter(c=>!seen.has(c.id)).map(c=>({id:c.id,path:c.path,sourceSha256:c.sourceSha256}));
 return {schemaVersion:1,consumer:'bun',scope:'Historical staging identity comparison only, separate from canonical mappings and runtime outcomes',executionCredit:false,nativeDeclarations:cases.length,stagingRows:mappings.length,identitiesPresent:mappings.filter(m=>m.state==='identity-present-needs-review').length,identitiesMissing:mappings.filter(m=>m.state==='identity-missing').length,withoutStagingCount:withoutStaging.length,stagingSources:files.map(f=>({path:f.path,sha256:hash(f.source)})).sort((a,b)=>a.path.localeCompare(b.path)),mappings,withoutStaging};
}
export async function buildReconciliation(root=resolve(import.meta.dir,'..')){
 const native=await inventoryNativeTests(root);if(native.unresolved.length)throw Error('Unresolved native registrations');
 const files=[];for await(const path of new Bun.Glob('docs/behaviors/staging-*.json').scan({cwd:root}))if(path!=='docs/behaviors/staging-reconciliation.json')files.push({path,source:await Bun.file(join(root,path)).text()});
 return reconcile(native.cases,files);
}
export async function checkReconciliation(root=resolve(import.meta.dir,'..')){
 const path=join(root,'docs/behaviors/staging-reconciliation.json'),expected=JSON.stringify(await buildReconciliation(root),null,2)+'\n';
 if(!await Bun.file(path).exists()||await Bun.file(path).text()!==expected)throw Error('Staging reconciliation stale; run bun scripts/mapping-reconciliation.ts');
}
if(import.meta.main){const root=resolve(import.meta.dir,'..');if(process.argv.includes('--check'))await checkReconciliation(root);else await Bun.write(join(root,'docs/behaviors/staging-reconciliation.json'),JSON.stringify(await buildReconciliation(root),null,2)+'\n');const report=await buildReconciliation(root);console.log(`${report.stagingRows} staging IDs: ${report.identitiesPresent} present, ${report.identitiesMissing} missing; ${report.withoutStagingCount} declarations without historical staging; no execution credit`);}
