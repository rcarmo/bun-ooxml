import {OoxmlError} from '../errors.ts';
import {type OpcPackage,type Relationship} from '../opc/package.ts';
import {parseXml,elements,attribute,type XmlElement} from '../xml/index.ts';

const P='http://schemas.openxmlformats.org/presentationml/2006/main';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const C='http://schemas.openxmlformats.org/package/2006/content-types';
const IMAGE=R+'/image';
function fail(message:string):never {throw new OoxmlError('PPTX_PICTURE_INVALID',message);}
const is=(node:XmlElement,name:string,namespace=P)=>node.localName===name&&node.namespaceURI===namespace;
function one(parent:XmlElement,name:string,namespace=P):XmlElement {
  const rows=parent.children.filter(node=>is(node,name,namespace));
  if(rows.length!==1)fail(`Expected exactly one ${name}`);
  return rows[0]!;
}
function optional(parent:XmlElement,name:string,namespace=A):XmlElement|undefined {
  const rows=parent.children.filter(node=>is(node,name,namespace));
  if(rows.length>1)fail(`Ambiguous ${name}`);
  return rows[0];
}
function integer(raw:string|undefined,label:string,min:number,max:number):number {
  if(raw===undefined||!/^[-+]?\d+$/.test(raw))fail(`Missing or invalid ${label}`);
  const value=Number(raw);
  if(!Number.isSafeInteger(value)||value<min||value>max)fail(`Out-of-range ${label}`);
  return value;
}
function flag(raw:string|undefined,label:string):boolean {
  if(raw===undefined||raw==='0'||raw==='false')return false;
  if(raw==='1'||raw==='true')return true;
  return fail(`Invalid ${label}`);
}

/** Direct DrawingML coordinates only. No layout inheritance or group-space conversion. */
export type PictureTransform={x:number;y:number;width:number;height:number;rotation:number;flipH:boolean;flipV:boolean};
export type PictureGroupTransform=PictureTransform&{childX:number;childY:number;childWidth:number;childHeight:number};
export type PictureGroup={shapeId:number;name:string;transform:PictureGroupTransform|null};
export type PictureEmbeddedAsset={relationshipId:string;target:string;partName:string;contentType:string;byteLength:number};
export type PictureLinkedAsset={relationshipId:string;target:string};
export type PictureCrop={left:number;top:number;right:number;bottom:number};
export type PictureInfo={
  shapeId:number;name:string;description:string|null;slidePart:string;groups:PictureGroup[];
  embedded:PictureEmbeddedAsset|null;linked:PictureLinkedAsset|null;
  transform:PictureTransform|null;crop:PictureCrop;
};
function identity(parent:XmlElement,nvName:string):{shapeId:number;name:string;description:string|null} {
  const properties=one(one(parent,nvName),'cNvPr');
  return {
    shapeId:integer(attribute(properties,'id'),'shape ID',1,2147483647),
    name:attribute(properties,'name')??'',description:attribute(properties,'descr')??null,
  };
}
const coordinate=(node:XmlElement,key:string)=>integer(attribute(node,key),key,-Number.MAX_SAFE_INTEGER,Number.MAX_SAFE_INTEGER);
const extent=(node:XmlElement,key:string)=>integer(attribute(node,key),key,0,Number.MAX_SAFE_INTEGER);
function transform(parent:XmlElement,group:false):PictureTransform|null;
function transform(parent:XmlElement,group:true):PictureGroupTransform|null;
function transform(parent:XmlElement,group:boolean):PictureTransform|PictureGroupTransform|null {
  const x=optional(parent,'xfrm');
  if(!x)return null;
  const off=one(x,'off',A),size=one(x,'ext',A);
  const value:PictureTransform={x:coordinate(off,'x'),y:coordinate(off,'y'),width:extent(size,'cx'),height:extent(size,'cy'),rotation:integer(attribute(x,'rot')??'0','rotation',-2147483648,2147483647),flipH:flag(attribute(x,'flipH'),'flipH'),flipV:flag(attribute(x,'flipV'),'flipV')};
  if(!group)return value;
  const childOff=one(x,'chOff',A),childSize=one(x,'chExt',A);
  return {...value,childX:coordinate(childOff,'x'),childY:coordinate(childOff,'y'),childWidth:extent(childSize,'cx'),childHeight:extent(childSize,'cy')};
}
function crop(fill:XmlElement):PictureCrop {
  const rect=optional(fill,'srcRect');
  const value=(key:string)=>integer(rect?attribute(rect,key)??'0':'0','crop '+key,-2147483648,2147483647);
  return {left:value('l'),top:value('t'),right:value('r'),bottom:value('b')};
}
function imageContentType(pkg:OpcPackage,part:string):string {
  const root=parseXml(pkg.text('[Content_Types].xml')).root;
  if(!is(root,'Types',C))fail('Invalid content-type registry');
  const overrides=root.children.filter(n=>is(n,'Override',C)&&attribute(n,'PartName')==='/'+part);
  const extension=part.split('.').at(-1)!.toLowerCase();
  const defaults=root.children.filter(n=>is(n,'Default',C)&&attribute(n,'Extension')?.toLowerCase()===extension);
  if(overrides.length>1||defaults.length>1)fail('Ambiguous image content type');
  const type=attribute(overrides[0]??defaults[0]??root,'ContentType');
  if(!type||!/^image\/[^\s/]+$/.test(type))fail(`Non-image content type for ${part}`);
  return type;
}
function relation(relationships:Relationship[],id:string):Relationship {
  const matches=relationships.filter(r=>r.id===id);
  if(matches.length!==1||matches[0]!.type!==IMAGE)fail(`Missing or wrong image relationship ${id}`);
  return matches[0]!;
}
function assets(pkg:OpcPackage,part:string,fill:XmlElement):Pick<PictureInfo,'embedded'|'linked'> {
  const blip=one(fill,'blip',A);
  // An unqualified or foreign-namespace lookalike must not silently erase an asset.
  for(const name of Object.keys(blip.attributes)) {
    if(['embed','link'].includes(name.split(':').at(-1)!)&&blip.attributeNamespaces[name]!==R)fail('Image relationship attribute has the wrong namespace');
  }
  const embed=attribute(blip,'embed',R),link=attribute(blip,'link',R);
  if(!embed&&!link)fail('Picture has no image relationship');
  const relationships=pkg.relationships(part);
  let embedded:PictureEmbeddedAsset|null=null,linked:PictureLinkedAsset|null=null;
  if(embed){
    const r=relation(relationships,embed);
    if(r.external||!r.resolved)fail('Embedded image must target an internal part');
    const data=pkg.get(r.resolved);
    if(!data)fail('Embedded image part is missing');
    embedded={relationshipId:embed,target:r.target,partName:r.resolved,contentType:imageContentType(pkg,r.resolved),byteLength:data.length};
  }
  if(link){
    const r=relation(relationships,link);
    if(!r.external)fail('Linked picture requires an external relationship');
    // Never fetch, stat, parse, or canonicalise an external target.
    linked={relationshipId:link,target:r.target};
  }
  return {embedded,linked};
}

/** Detached source-order picture metadata. Unknown payload/effect XML is retained, never rewritten. */
export function inspectPictures(pkg:OpcPackage,part:string):PictureInfo[] {
  const doc=parseXml(pkg.text(part));
  if(!is(doc.root,'sld'))fail('Expected a presentation slide');
  const tree=one(one(doc.root,'cSld'),'spTree'),ids=new Set<number>();
  for(const node of elements(doc,'cNvPr',P)){
    const id=integer(attribute(node,'id'),'shape ID',1,2147483647);
    if(ids.has(id))fail(`Duplicate shape identity ${id}`);
    ids.add(id);
  }
  const result:PictureInfo[]=[],visited=new Set<XmlElement>();
  function visit(parent:XmlElement,groups:PictureGroup[]):void {
    for(const node of parent.children){
      if(is(node,'pic')){
        visited.add(node);
        const id=identity(node,'nvPicPr'),fill=one(node,'blipFill'),properties=one(node,'spPr');
        result.push({...id,slidePart:part,groups:groups.map(g=>({...g,transform:g.transform?{...g.transform}:null})),...assets(pkg,part,fill),transform:transform(properties,false),crop:crop(fill)});
      }else if(is(node,'grpSp')){
        const id=identity(node,'nvGrpSpPr'),properties=one(node,'grpSpPr');
        visit(node,[...groups,{shapeId:id.shapeId,name:id.name,transform:transform(properties,true)}]);
      }
    }
  }
  visit(tree,[]);
  if(elements(doc,'pic',P).some(p=>!visited.has(p)))fail('Picture is outside a direct shape tree/group owner');
  return result;
}
