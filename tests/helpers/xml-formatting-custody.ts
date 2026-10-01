import assert from 'node:assert/strict';
/** Independent lexical oracle: tokens are consumed, never searched inside quoted values. */
export function omitOpeningAttributes(xml:string,selected:ReadonlySet<string>):string {
 let at=0,result='';while(at<xml.length){const start=xml.indexOf('<',at);if(start<0){result+=xml.slice(at);break;}result+=xml.slice(at,start);const head=xml.slice(start).match(/^<[^\s/>]+/);if(!head||/^<\//.test(head[0])){const end=xml.indexOf('>',start);assert(end>=start);result+=xml.slice(start,end+1);at=end+1;continue;}
 let pos=start+head[0].length;result+=head[0];for(;;){const rest=xml.slice(pos),end=rest.match(/^\s*\/?\s*>/);if(end){result+=end[0];pos+=end[0].length;break;}const token=rest.match(/^(\s+)([^\s=/>]+)(\s*=\s*)(["'])([^]*?)\4/);assert(token,'invalid lexical oracle token');if(!selected.has(token[2]!))result+=token[0];pos+=token[0].length;}at=pos;}
 return result;
}
