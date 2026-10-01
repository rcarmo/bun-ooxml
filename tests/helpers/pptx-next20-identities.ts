import assert from 'node:assert/strict';
import sharedPin from '../../docs/behaviors/pptx-manipulation-candidate.json';
export const pptxNext20Kinds=['patch-title','patch-body','patch-subtitle','append-title','bullet-default','bullet-sequence','bullet-level','bullet-bold-label','clear-bullets','autofit-shrink','autofit-none','autofit-resize','insert-start','insert-middle','reorder','reorder-refusal','table-values','table-geometry','set-notes','notes-readback'];
export const pptxNext20Ids=pptxNext20Kinds.map(k=>'@id-bun-pptx-next20-'+k);
export function withoutPptxNext20(keys:string[]):string[]{
 const shared=keys.filter(k=>k.startsWith('@id-pptx-manipulation-'));if(shared.length){assert.equal(shared.length,20);assert.deepEqual([...shared].sort(),[...sharedPin.selectedScenarioIds].sort());keys=keys.filter(k=>!shared.includes(k));}
 const added=keys.filter(k=>k.startsWith('@id-bun-pptx-next20-'));if(added.length===0)return keys;assert.equal(added.length,20);assert.deepEqual([...added].sort(),[...pptxNext20Ids].sort());return keys.filter(k=>!added.includes(k));
}
