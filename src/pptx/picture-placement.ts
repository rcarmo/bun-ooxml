import {OoxmlError} from '../errors.ts';
import {pictureRequest,type PictureGeometry,type PictureOptions,type PictureReceipt} from './picture-add.ts';
import type {PictureCrop} from './pictures.ts';
export type PictureFit='contain'|'cover'|'stretch';
export type PictureIntrinsicSize={width:number;height:number};
export type PicturePlacement={geometry:PictureGeometry;crop:PictureCrop};
export type FittedPictureOptions=PictureOptions&{fit:PictureFit;intrinsic:PictureIntrinsicSize};
export type FittedPictureReceipt=PictureReceipt&PicturePlacement;
function fail(message:string):never {throw new OoxmlError('PPTX_PICTURE_UNSUPPORTED',message);}
function data(value:unknown,keys:string[]):Record<string,unknown>{
 if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Fit inputs require plain data');
 const copy:Record<string,unknown>={};for(const key of Reflect.ownKeys(value)){if(typeof key!=='string'||!keys.includes(key))fail('Unknown fit property');const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Fit accessors are unsupported');copy[key]=d.value;}return copy;
}
function bounded(value:unknown,min:number):number {if(typeof value!=='number'||!Number.isInteger(value)||value<min||value>2147483647)fail('Fit dimensions/coordinates must be bounded integers');return value;}
/** Exact rational calculations; explicit pixel ratio, no image decoding or DPI inference. */
export function calculatePicturePlacement(box:PictureGeometry,intrinsic:PictureIntrinsicSize,fit:PictureFit):PicturePlacement{
 const b=data(box,['x','y','width','height']),i=data(intrinsic,['width','height']);
 const x=bounded(b.x,-2147483648),y=bounded(b.y,-2147483648),width=bounded(b.width,1),height=bounded(b.height,1),iw=BigInt(bounded(i.width,1)),ih=BigInt(bounded(i.height,1)),bw=BigInt(width),bh=BigInt(height);
 if(!['contain','cover','stretch'].includes(fit))fail('Unknown picture fit policy');
 const geometry={x,y,width,height},crop={left:0,top:0,right:0,bottom:0},wide=iw*bh>ih*bw;
 if(fit==='contain'){
  if(wide)geometry.height=Number(bw*ih/iw);else geometry.width=Number(bh*iw/ih);
  if(geometry.width<1||geometry.height<1)fail('Contained extent rounds below one EMU');
  geometry.x=bounded(x+Math.floor((width-geometry.width)/2),-2147483648);geometry.y=bounded(y+Math.floor((height-geometry.height)/2),-2147483648);
 }else if(fit==='cover'){
  const denominator=wide?iw*bh:ih*bw,numerator=wide?iw*bh-ih*bw:ih*bw-iw*bh;
  // nearest integer, ties upward; each side discards half the fraction.
  const side=Number((numerator*100000n+denominator)/(2n*denominator));
  if(side*2>=100000)fail('Cover crop rounds to an empty visible region');
  if(wide)crop.left=crop.right=side;else crop.top=crop.bottom=side;
 }
 return {geometry,crop};
}
export function fittedPictureRequest(bytes:Uint8Array,box:PictureGeometry,options:FittedPictureOptions){
 const o=data(options,['contentType','name','description','fit','intrinsic']);
 const placement=calculatePicturePlacement(box,o.intrinsic as PictureIntrinsicSize,o.fit as PictureFit),request=pictureRequest(bytes,placement.geometry,{contentType:o.contentType as PictureOptions['contentType'],...(o.name!==undefined?{name:o.name as string}:{}),...(o.description!==undefined?{description:o.description as string}:{})});
 return {request,placement};
}
