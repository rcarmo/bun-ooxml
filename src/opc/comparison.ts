import {xmlEquivalent} from '../xml/comparison.ts';
import {admitPackage} from './admission.ts';
import {sameBytes} from './package.ts';
import type {ZipLimits} from './zip.ts';

/** File payload categories. XML equivalence is the bounded xmlEquivalent profile,
 * not a schema, relationship-graph, rendering or signature guarantee. */
export interface PackageComparison {
 added:string[];
 removed:string[];
 changed:string[];
 equivalent_xml:string[];
 unchanged:string[];
}

/** Compare admitted archive members without modifying inputs or requiring OPC.
 * Both archives independently use the same resource limits. Admission failure
 * throws before any report is returned, even for equal or one-sided XML members.
 * ZIP compression, metadata and empty directory entries are not payload changes.
 */
export function comparePackageArchives(before:Uint8Array,after:Uint8Array,limits:ZipLimits={}):PackageComparison {
 const left=admitPackage(before,limits),right=admitPackage(after,limits);
 const report:PackageComparison={added:[],removed:[],changed:[],equivalent_xml:[],unchanged:[]};
 const names=[...new Set([...left.keys(),...right.keys()])].sort();
 for(const name of names){
  const a=left.get(name),b=right.get(name);
  if(a===undefined){report.added.push(name);continue;}
  if(b===undefined){report.removed.push(name);continue;}
  if(sameBytes(a,b)){report.unchanged.push(name);continue;}
  if(/\.(?:xml|rels)$/i.test(name)&&xmlEquivalent(a,b))report.equivalent_xml.push(name);
  else report.changed.push(name);
 }
 return report;
}
