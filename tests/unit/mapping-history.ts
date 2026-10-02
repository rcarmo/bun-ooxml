import cellMigration from '../../docs/behaviors/cell-runtime-mapping-migration.json';
import uniformApi18 from '../../docs/behaviors/uniform-api18-candidate.json';
import contract20 from '../../docs/behaviors/contract20-candidate.json';
// Historical mapping fingerprints precede the planned-case expansion and
// negative-budget and descriptor bindings. Restore only the reviewed global
// inventory assertions and two reclassified partial associations; keep every
// other row intact.
const admissionId='bun:tests/unit/package-admission.test.ts:invalid admission limit configuration refuses rather than disabling bounds';
const descriptorId='bun:tests/unit/zip.test.ts:readZip / reads a valid archive with a directory entry, data descriptor, and declared comment';
export function beforeUniformApi18MappingRows<T>(rows:T[]):T[]{return rows.map(input=>{
 return input&&typeof input==='object'&&'caseKeys' in input?{...input,caseKeys:(input as any).caseKeys.map((key:string)=>contract20.caseIdentityMigration.find(r=>r.after===key)?.before??key).map((key:string)=>uniformApi18.caseIdentityMigration.find(r=>r.after===key)?.before??key)} as T:input;
});}
export function beforeSharedPlannedExpansion<T extends {testId:string;assertions:string[];scenarioIds:string[];gaps:string[]}>(rows:T[]):T[]{
 return beforeUniformApi18MappingRows(rows).map(current=>{
  const change=cellMigration.changes.find(r=>r.testId===current.testId);
  let row=current;
  if(change){
   for(const key of ['assertions','outcomes','gaps','scenarioIds','caseKeys'] as const)if(JSON.stringify((current as any)[key])!==JSON.stringify(change.after[key]))throw Error('Unreviewed cell runtime mapping: '+key);
   row={...current,assertions:change.before.assertions,...('outcomes' in current?{outcomes:change.before.outcomes}:{}),gaps:change.before.gaps} as T;
  }
  const restored={...row,assertions:row.assertions.map(a=>a.replace('expect(inv.counts.cases.planned).toBe(59)','expect(inv.counts.cases.planned).toBe(46)').replace('expect(inv.counts.cases.implemented).toBe(732)','expect(inv.counts.cases.implemented).toBe(726)'))};
  if(row.testId===admissionId){
   if(JSON.stringify(row.scenarioIds)!==JSON.stringify(['@id-package-admission-negative-budget']))throw Error('Admission source association drift');
   const added=row.gaps.filter(g=>g.includes('two sealed DOCX rows'));
   if(added.length!==1)throw Error('Admission gap custody drift');
   return {...restored,scenarioIds:['@id-package-admission-resource-limits'],gaps:row.gaps.filter(g=>g!==added[0])};
  }
  if(row.testId===descriptorId){
   const added=row.gaps.filter(g=>g.includes('unsigned descriptor whose CRC bytes collide'));
   if(added.length!==1)throw Error('Descriptor gap custody drift');
   return {...restored,gaps:row.gaps.filter(g=>g!==added[0])};
  }
  return restored;
 });
}
