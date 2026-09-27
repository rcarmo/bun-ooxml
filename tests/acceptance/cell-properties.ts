import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const document=(c:Record<string,unknown>)=>(c.state as {document:Document}).document;
const cell=(c:Record<string,unknown>)=>document(c).tables[0]!.cell(0,0);
export const scenarioIds=['@id-docx-go-cell-shading-getter','@id-docx-go-cell-properties-getters'];
export const bindings:StepBinding[]=[
 {pattern:/^cell zero-zero of a new Go Word (two-by-two|one-by-one) table$/,run:(c,size)=>{const d=Document.create(),n=size==='two-by-two'?2:1;d.addTable(n,n);(c.state as {document:Document}).document=d;}},
 {pattern:/^its shading is set to (\S+)$/,run:(c,color)=>{cell(c).setProperties({shading:color!});}},
 {pattern:/^its shading getter equals (\S+)$/,run:(c,color)=>{assert.equal(cell(c).directProperties().shading,color);}},
 {pattern:/^width is set to 2400 dxa, vertical alignment center and text direction tbRl$/,run:c=>{cell(c).setProperties({widthTwips:2400,verticalAlign:'center',textDirection:'tbRl'});}},
 {pattern:/^a top border with single style, size eight and colour 000000 is assigned$/,run:c=>{cell(c).setProperties({topBorder:{style:'single',size:8,color:'000000'}});}},
 {pattern:/^width equals (\d+), width type (\S+), alignment (\S+) and direction (\S+)$/,run:(c,width,type,alignment,direction)=>{
  const v=cell(c).directProperties();assert.equal(v.widthTwips,Number(width));assert.equal(v.verticalAlign,alignment);assert.equal(v.textDirection,direction);
  const xml=parseXml(new TextDecoder().decode(document(c).package.get('word/document.xml'))),tcW=elements(xml,'tcW',W)[0];assert(tcW);assert.equal(attribute(tcW,'type',W),type);
 }},
 {pattern:/^the border collection and top border are nonnil$/,run:c=>{
  const xml=parseXml(new TextDecoder().decode(document(c).package.get('word/document.xml'))),borders=elements(xml,'tcBorders',W);assert.equal(borders.length,1);assert.equal(borders[0]!.children.filter(n=>n.namespaceURI===W&&n.localName==='top').length,1);
  assert.deepEqual(cell(c).directProperties().topBorder,{style:'single',size:8,color:'000000'});
 }},
];
