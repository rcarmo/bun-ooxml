import {OoxmlError} from '../errors.ts';
import type {BodyAnchor} from './index.ts';
export interface BodyPlaceholder {
 readonly bodyIndex:number;
 /** Offsets into the decoded paragraph text, in UTF-16 code units. */
 readonly start:number;
 readonly end:number;
 readonly text:string;
 readonly name:string;
}
export interface BodyMap {
 /** Sections counts direct-outline heading markers, not Word section breaks. */
 readonly counts:Readonly<{sections:number;tables:number;placeholders:number;anchors:number}>;
 readonly anchors:readonly BodyAnchor[];
 readonly placeholders:readonly BodyPlaceholder[];
}
/** Linear scan of nonnested angle-delimited text; no regex backtracking or cross-paragraph matches. */
export function bodyMap(anchors:readonly BodyAnchor[],tables:number):BodyMap {
 const placeholders:BodyPlaceholder[]=[];
 let sections=0;
 for(const anchor of anchors){
  if(anchor.type==='section_heading')sections++;
  const text=anchor.text;let start=-1,depth=0,nested=false;
  for(let at=0;at<text.length;at++){
   if(text[at]==='<'){if(depth===0){start=at;nested=false;}else nested=true;depth++;}
   else if(text[at]==='>'&&depth>0){
    if(--depth!==0)continue;
    if(nested||at-start<=1||at-start-1>256)continue;
    const name=text.slice(start+1,at);if(!name.trim()||/[\r\n\t]/.test(name))continue;
    if(placeholders.length===10000)throw new OoxmlError('docx-body-map-limit','At most 10000 body placeholder occurrences are supported');
    placeholders.push(Object.freeze({bodyIndex:anchor.bodyIndex,start,end:at+1,text:text.slice(start,at+1),name}));
   }
  }
 }
 return Object.freeze({counts:Object.freeze({sections,tables,placeholders:placeholders.length,anchors:anchors.length}),anchors,placeholders:Object.freeze(placeholders)});
}
