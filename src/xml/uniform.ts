import {UniformApiError,profileOperation,type UniformApiCategory} from '../uniform.ts';
import {OoxmlError} from '../errors.ts';
import {XmlSnapshot,type XmlRemovalTarget,type XmlAttributePatch,type XmlStructurePatch} from './removal.ts';
const MAPPING:Readonly<Record<string,UniformApiCategory>>={
 XML_REMOVAL_ROOT:'root',XML_STRUCTURE_ROOT:'root',
 XML_REMOVAL_OVERLAP:'overlap-or-duplicate',XML_STRUCTURE_OVERLAP:'overlap-or-duplicate',XML_ATTRIBUTE_DUPLICATE:'overlap-or-duplicate',XML_STRUCTURE_ATTRIBUTE:'overlap-or-duplicate',
 XML_REMOVAL_TARGET:'foreign-target',XML_STRUCTURE_TARGET:'foreign-target',XML_ATTRIBUTE_TARGET:'foreign-target',
 XML_REMOVAL_SOURCE:'invalid-name-or-value',XML_ATTRIBUTE_PATCH:'invalid-name-or-value',XML_ATTRIBUTE_NAME:'invalid-name-or-value',XML_STRUCTURE_INPUT:'invalid-name-or-value',XML_STRUCTURE_NAME:'invalid-name-or-value',XML_STRUCTURE_CYCLE:'invalid-name-or-value',
 XML_INPUT_TOO_LARGE:'xml-edit-limit',XML_DEPTH_LIMIT:'xml-edit-limit',XML_NODE_LIMIT:'xml-edit-limit',XML_ATTRIBUTE_LIMIT:'xml-edit-limit',XML_STRUCTURE_LIMIT:'xml-edit-limit',
 XML_MALFORMED:'unsafe-XML',XML_MISMATCHED_TAG:'unsafe-XML',XML_INVALID_CHAR:'unsafe-XML',XML_DTD_FORBIDDEN:'unsafe-XML',XML_ENTITY_FORBIDDEN:'unsafe-XML',XML_DUPLICATE_ATTRIBUTE:'unsafe-XML',XML_UNBOUND_PREFIX:'unsafe-XML',XML_REMOVAL_UNSAFE:'unsafe-XML',XML_STRUCTURE_UNSAFE:'unsafe-XML',XML_ATTRIBUTE_UNSAFE:'unsafe-XML',
};
function xmlOperation<T>(operation:()=>T):T{
 return profileOperation(()=>{try{return operation();}catch(error){
  if(error instanceof OoxmlError&&['XML_ATTRIBUTE_UNSAFE','XML_STRUCTURE_UNSAFE','XML_REMOVAL_UNSAFE'].includes(error.code)&&error.cause instanceof OoxmlError&&['XML_DEPTH_LIMIT','XML_NODE_LIMIT','XML_INPUT_TOO_LARGE'].includes(error.cause.code))throw new UniformApiError('xml-edit-limit',error.message,{cause:error});
  throw error;
 }},MAPPING);
}
/** Additive Unicode/no-BOM profile; all handles belong to an immutable snapshot. */
export class UniformXmlSnapshot {
 readonly elements:readonly XmlRemovalTarget[];
 readonly #snapshot:XmlSnapshot;
 readonly #owned:ReadonlySet<XmlRemovalTarget>;
 private constructor(snapshot:XmlSnapshot){this.#snapshot=snapshot;this.elements=snapshot.elements;this.#owned=new Set(this.elements);Object.freeze(this);}
 static parse(source:string):UniformXmlSnapshot{
  if(typeof source!=='string'||source.startsWith('\ufeff'))throw new UniformApiError('invalid-name-or-value','Expected Unicode XML without a leading BOM');
  return xmlOperation(()=>new UniformXmlSnapshot(XmlSnapshot.parse(source)));
 }
 private batch(value:unknown):void{
  if(!Array.isArray(value))throw new UniformApiError('invalid-name-or-value','Expected a batch array');
  if(value.length>100000)throw new UniformApiError('xml-edit-limit','Batch exceeds 100000 patches');
 }
 setAttributes(patches:readonly XmlAttributePatch[]):string{this.batch(patches);return xmlOperation(()=>this.#snapshot.setAttributes(patches));}
 appendChildren(patches:readonly XmlStructurePatch[]):string{this.batch(patches);return xmlOperation(()=>this.#snapshot.appendChildren(patches));}
 replaceElements(patches:readonly XmlStructurePatch[]):string{this.batch(patches);return xmlOperation(()=>this.#snapshot.replaceElements(patches));}
 remove(targets:readonly XmlRemovalTarget[]):string{
  this.batch(targets);
  // Legacy removal limits arrays by node count before checking duplicates. The
  // profile checks issuance first so repeated owned handles have a distinct refusal.
  for(let i=0;i<targets.length;i++){const t=targets[i];if(!t||!this.#owned.has(t))throw new UniformApiError('foreign-target','Target was not issued by this snapshot');if(t===this.elements[0])throw new UniformApiError('root','Root cannot be removed');}
  const seen=new Set<XmlRemovalTarget>();for(let i=0;i<targets.length;i++){const t=targets[i]!;if(seen.has(t))throw new UniformApiError('overlap-or-duplicate','Duplicate target');seen.add(t);}
  return xmlOperation(()=>this.#snapshot.remove(targets));
 }
}
