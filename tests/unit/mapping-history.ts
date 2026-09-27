// Historical mapping fingerprints exclude the new shared planned cases.
// Restore only the former global planned-count assertion before comparing
// unchanged prior rows; all other fields and native assertions remain intact.
export function beforeSharedPlannedExpansion<T extends {assertions:string[]}>(rows:T[]):T[]{
 return rows.map(row=>({...row,assertions:row.assertions.map(a=>a.replace('expect(inv.counts.cases.planned).toBe(59)','expect(inv.counts.cases.planned).toBe(46)'))}));
}
