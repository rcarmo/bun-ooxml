import assert from 'node:assert/strict';
/** @script Author a practical four-slide graphics sample for external application checks.
 * @usage bun scripts/graphics-sample.ts [output-directory]
 * @description Uses centrally sealed candidate recipes; does not confer Office interoperability.
 */
import {resolve} from 'node:path';
import {mkdir} from 'node:fs/promises';
import {Presentation} from '../src/pptx/index.ts';
import {insertionCases,insertionInput} from '../tests/helpers/picture-insertion-inputs.ts';
const root=resolve(process.argv[2]??resolve(import.meta.dir,'../artifacts/graphics-uno')),p=Presentation.create();
await mkdir(root,{recursive:true});
const png=(await insertionInput(insertionCases.find((c:any)=>c.caseId==='png'))).source;
const jpeg=(await insertionInput(insertionCases.find((c:any)=>c.caseId==='jpeg'))).source;
const inch=914400;const box=(x:number,y:number,w:number,h:number)=>({x:Math.round(x*inch),y:Math.round(y*inch),width:Math.round(w*inch),height:Math.round(h*inch)});
let s=p.addTextSlide('Raster, crop, transforms and SVG');
const first=s.addPicture(png,box(.5,1.5,2,1.5),{contentType:'image/png'});s.setPictureCrop(first.shapeId,{left:10000,top:5000,right:10000,bottom:5000});s.patchPictureTransform(first.shapeId,{rotation:600000,flipH:true});s.setPictureTransparency(first.shapeId,20000);
s.addFittedPicture(jpeg,box(3,1.5,2,1.5),{contentType:'image/jpeg',fit:'contain',intrinsic:{width:640,height:480}});
s.addSvgPicture(new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100"><rect x="5" y="5" width="90" height="90" fill="#3b82f6"/></svg>'),png,box(5.5,1.5,2,1.5),{contentType:'image/png'});
const replaced=s.addPicture(png,box(.5,3.5,2,1.5),{contentType:'image/png'});s.replacePicture(replaced.shapeId,jpeg,{contentType:'image/jpeg'});
const removed=s.addPicture(png,box(8,5,1,1),{contentType:'image/png'});s.deletePicture(removed.shapeId);
s.addTextBox('Raster rotation / crop / alpha; contain-fit JPEG; SVG with fallback; isolated replacement',box(.5,5.5,8,1));
s=p.addTextSlide('AutoShapes, gradients, opacity and grouping');
const a=s.addAutoShape('roundRect',box(.5,1.5,2,1.25),{text:'Round rectangle',fill:'60A5FA',adjustments:{adj:25000}}),b=s.addAutoShape('ellipse',box(3,1.5,2,1.25),{text:'Ellipse',fill:'FBBF24'}),c=s.addAutoShape('diamond',box(5.5,1.5,2,1.25),{text:'Diamond',fill:'A78BFA'});
s.setLinearGradient(a.shapeId,{angle:5400000,scaled:true,stops:[{position:0,color:{kind:'srgb',value:'60A5FA'}},{position:100000,color:{kind:'srgb',value:'DBEAFE'}}]});s.setShapeOpacity(b.shapeId,65000);s.patchOutlineStyle(c.shapeId,{cap:'rnd',compound:'dbl',join:{type:'round'}});
const g1=s.addAutoShape('rect',box(1,3.5,1.5,1),{text:'Grouped A',fill:'86EFAC'}),g2=s.addAutoShape('triangle',box(3,3.5,1.5,1),{text:'Grouped B',fill:'FCA5A5'});s.groupShapes([g1.shapeId,g2.shapeId],box(.5,3,5,2));
s=p.addTextSlide('Editable diagrams and attached connectors');
const diagram=s.addDiagram([{key:'input',text:'Input'},{key:'process',text:'Process'},{key:'output',text:'Output'}],[{from:'input',to:'process'},{from:'process',to:'output'}],{x:Math.round(.5*inch),y:2*inch,nodeWidth:2*inch,nodeHeight:inch,gap:inch/2,direction:'row'});
for(const e of diagram.edges)s.patchOutlineStyle(e.shapeId,{headEnd:{type:'triangle',width:'med',length:'med'},cap:'rnd'});
s.addTextBox('Three editable nodes connected by exact shape IDs and connection sites.',box(.5,4,8,1));
s=p.addTextSlide('Freeform vector geometry and z-order');
const free=s.addFreeform(box(1,2,3,2),{width:300,height:200,commands:[{op:'move',x:0,y:200},{op:'line',x:150,y:0},{op:'line',x:300,y:200},{op:'close'}]},{fill:'34D399',lineColor:'065F46',lineWidth:25400});const over=s.addAutoShape('ellipse',box(2.5,2.5,2,1.5),{fill:'F472B6',text:'Front'});s.reorderShapes([free.shapeId,over.shapeId]);
const path=root+'/graphics-showcase.pptx';await p.save(path);
const reopened=await Presentation.open(path);assert.equal(reopened.slides.length,4);assert.equal(reopened.slides[0]!.inspectPictures().length,4);
console.log('Saved practical four-slide graphics sample: '+path);
