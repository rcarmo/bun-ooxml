import {expect} from 'bun:test';
import {parseXml,attribute,type XmlDocument} from '../../src/xml/index.ts';
import {OoxmlError} from '../../src/errors.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const bindings:StepBinding[]=[
 {pattern:/^XML with a numeric-leading (element-local|attribute-local|declared-prefix) QName component$/,run:(ctx,kind)=>{const inputs:Record<string,string>={'element-local':'<p:1 xmlns:p="urn:test"/>','attribute-local':'<r xmlns:p="urn:test" p:1="x"/>','declared-prefix':'<r xmlns:1="urn:test"/>'};ctx.source=inputs[kind];}},
 {pattern:/^the namespace-name fixture is parsed$/,run:ctx=>{try{ctx.doc=parseXml(ctx.source as string);}catch(error){ctx.error=error;}}},
 {pattern:/^parsing refuses with the malformed XML code$/,run:ctx=>{expect(ctx.error).toBeInstanceOf(OoxmlError);expect((ctx.error as OoxmlError).code).toBe('XML_MALFORMED');expect(ctx.doc).toBeUndefined();}},
 {pattern:/^XML with valid Unicode prefix and local name components$/,run:ctx=>{ctx.source='<π:名 xmlns:π="urn:unicode" π:é="value"><π:𐐀/></π:名>';}},
 {pattern:/^expanded element and attribute names retain their Unicode identity$/,run:ctx=>{expect(ctx.error).toBeUndefined();const d=ctx.doc as XmlDocument;expect(d.root.localName).toBe('名');expect(d.root.namespaceURI).toBe('urn:unicode');expect(attribute(d.root,'é','urn:unicode')).toBe('value');expect(d.root.children[0]!.localName).toBe('𐐀');}},
 {pattern:/^source offsets still address the original Unicode element$/,run:ctx=>{const d=ctx.doc as XmlDocument;const child=d.root.children[0]!;expect((ctx.source as string).slice(child.start,child.end)).toBe('<π:𐐀/>');}},
 {pattern:/^XML with a non-breaking space (before|after) the root element$/,run:(ctx,position)=>{ctx.source=position==='before'?'\u00a0<r/>':'<r/>\u00a0';}},
];
