import {OoxmlError} from './errors.ts';
export type UniformApiCategory='root'|'overlap-or-duplicate'|'foreign-target'|'invalid-name-or-value'|'unsafe-XML'|'xml-edit-limit'|'unsupported-static-reference'|'static-reference-limit'|'unsupported-direct-range'|'invalid-reference-insertion';
/** Stable refusal of the additive uniform profile; legacy error codes stay available. */
export class UniformApiError extends OoxmlError {
 readonly category:UniformApiCategory;
 constructor(category:UniformApiCategory,message:string,options?:ErrorOptions){super(category,message,options);this.category=category;}
}
export type UniformApiResult<T>={ok:true;value:T}|{ok:false;category:UniformApiCategory;value:null};
/** Only typed profile refusals become data; unexpected implementation errors propagate. */
export function uniformApiResult<T>(operation:()=>T):UniformApiResult<T>{
 try{return {ok:true,value:operation()};}catch(error){if(error instanceof UniformApiError)return {ok:false,category:error.category,value:null};throw error;}
}
export function profileOperation<T>(operation:()=>T,mapping:Readonly<Record<string,UniformApiCategory>>):T{
 try{return operation();}catch(error){if(error instanceof UniformApiError)throw error;if(error instanceof OoxmlError){const category=mapping[error.code];if(category)throw new UniformApiError(category,error.message,{cause:error});}throw error;}
}
