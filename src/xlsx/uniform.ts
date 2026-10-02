import {UniformApiError,profileOperation,type UniformApiCategory} from '../uniform.ts';
import {analyzeProfileFormulaReferences,type FormulaReference} from './formula.ts';
import {parseA1Range,type A1Range} from './range.ts';
import {insertProfileFormulaReferences,type FormulaInsertion} from './formula-remap.ts';
const FORMULA:Readonly<Record<string,UniformApiCategory>>={'xlsx-formula-unsupported':'unsupported-static-reference','xlsx-formula-limit':'static-reference-limit'};
/** Detached exact records with UTF-8 byte spans, bounded grammar and sealed arities. */
export function analyzeStaticReferences(source:string):FormulaReference[]{return profileOperation(()=>analyzeProfileFormulaReferences(source),FORMULA);}
export function parseStaticRange(source:string):A1Range{
 if(typeof source==='string'&&/[\u0080-\u009f]/.test(source))throw new UniformApiError('unsupported-direct-range','C1 controls are unsupported');
 return profileOperation(()=>parseA1Range(source),{'xlsx-range-unsupported':'unsupported-direct-range'});
}
export function remapStaticReferences(source:string,contextSheet:string,change:FormulaInsertion):string{
 return profileOperation(()=>insertProfileFormulaReferences(source,contextSheet,change),{...FORMULA,'xlsx-formula-remap-unsupported':'invalid-reference-insertion'});
}
