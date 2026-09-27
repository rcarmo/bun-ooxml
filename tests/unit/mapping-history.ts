// Historical mapping fingerprints precede the planned-case expansion and
// negative-budget binding. Restore only the reviewed global inventory assertions
// and the one reclassified partial association; keep every other row intact.
const admissionId='bun:tests/unit/package-admission.test.ts:invalid admission limit configuration refuses rather than disabling bounds';
export function beforeSharedPlannedExpansion<T extends {testId:string;assertions:string[];scenarioIds:string[];gaps:string[]}>(rows:T[]):T[]{
 return rows.map(row=>{
  const restored={...row,assertions:row.assertions.map(a=>a.replace('expect(inv.counts.cases.planned).toBe(59)','expect(inv.counts.cases.planned).toBe(46)').replace('expect(inv.counts.cases.implemented).toBe(728)','expect(inv.counts.cases.implemented).toBe(726)'))};
  if(row.testId===admissionId){
   if(JSON.stringify(row.scenarioIds)!==JSON.stringify(['@id-package-admission-negative-budget']))throw Error('Admission source association drift');
   const added=row.gaps.filter(g=>g.includes('two sealed DOCX rows'));
   if(added.length!==1)throw Error('Admission gap custody drift');
   return {...restored,scenarioIds:['@id-package-admission-resource-limits'],gaps:row.gaps.filter(g=>g!==added[0])};
  }
  return restored;
 });
}
