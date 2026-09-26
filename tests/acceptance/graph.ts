import {expect} from 'bun:test';
import { fixturePath, F } from '../../scripts/fixture-inputs.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {OpcPackage} from '../../src/opc/package.ts';
import {addPart,removePart,addRelationship,removeRelationship} from '../../src/opc/graph.ts';
import {getContentType,setPartContentType} from '../../src/opc/content-types.ts';
import {diffPackages,type PackageDiffReport} from '../../src/opc/diff.ts';
import {OoxmlError} from '../../src/errors.ts';
const fixture=fixturePath(F.shared.presentPlaceholder);
const target='custom/data.bin';
async function prepare(ctx:Record<string,unknown>){const p=await OpcPackage.open(fixture);ctx.graph=p;ctx.original=await OpcPackage.open(p.toBytes());}
function attach(ctx:Record<string,unknown>){const p=ctx.graph as OpcPackage;addPart(p,target,new Uint8Array([7,8,9]),'application/octet-stream');ctx.rid=addRelationship(p,'','urn:test/data',target).id;}
export const bindings:StepBinding[]=[
 {pattern:/^the shared DOCX package for graph editing$/,run:prepare},
 {pattern:/^an opaque part and internal root relationship are added$/,run:attach},
 {pattern:/^the saved package reopens with the new part and its exact content type$/,run:async ctx=>{const p=await OpcPackage.open((ctx.graph as OpcPackage).toBytes());expect(p.get(target)).toEqual(new Uint8Array([7,8,9]));expect(getContentType(p,target)).toBe('application/octet-stream');expect(p.relationships().find(r=>r.id===ctx.rid)?.resolved).toBe(target);}},
 {pattern:/^unrelated original payloads retain their exact bytes$/,run:ctx=>{const a=ctx.original as OpcPackage,b=ctx.graph as OpcPackage;for(const name of a.names())if(!['[Content_Types].xml','_rels/.rels'].includes(name))expect(b.get(name)).toEqual(a.get(name)!);}},
 {pattern:/^the shared DOCX package with an additional related opaque part$/,run:async ctx=>{await prepare(ctx);attach(ctx);ctx.before=(ctx.graph as OpcPackage).toBytes();}},
 {pattern:/^removal of the referenced part is attempted$/,run:ctx=>{try{removePart(ctx.graph as OpcPackage,target);}catch(e){ctx.error=e;}}},
 {pattern:/^graph editing refuses with code "([^"]+)"$/,run:(ctx,code)=>{expect(ctx.error).toBeInstanceOf(OoxmlError);expect((ctx.error as OoxmlError).code).toBe(code);}},
 {pattern:/^the package bytes are unchanged after the refusal$/,run:ctx=>{expect((ctx.graph as OpcPackage).toBytes()).toEqual(ctx.before as Uint8Array);}},
 {pattern:/^its unreferenced root relationship and part are removed$/,run:ctx=>{const p=ctx.graph as OpcPackage;removeRelationship(p,'',ctx.rid as string);removePart(p,target);}},
 {pattern:/^the saved package no longer contains the part or its content-type override$/,run:async ctx=>{const p=await OpcPackage.open((ctx.graph as OpcPackage).toBytes());expect(p.get(target)).toBeUndefined();expect(p.text('[Content_Types].xml')).not.toContain('/'+target);}},
 {pattern:/^all retained relationship targets resolve$/,run:async ctx=>{await expect(OpcPackage.open((ctx.graph as OpcPackage).toBytes())).resolves.toBeInstanceOf(OpcPackage);}},
 {pattern:/^only a part content type is changed$/,run:ctx=>{const p=ctx.graph as OpcPackage;setPartContentType(p,'word/document.xml','application/vnd.test.document+xml');ctx.report=diffPackages(ctx.original as OpcPackage,p);}},
 {pattern:/^the package diff identifies the part despite identical payload hashes$/,run:ctx=>{const d=ctx.report as PackageDiffReport;expect(d.changed).toContain('word/document.xml');const part=d.parts.find(p=>p.name==='word/document.xml')!;expect(part.beforeSha256).toBe(part.afterSha256!);expect(part.beforeContentType).not.toBe(part.afterContentType!);}},
 {pattern:/^the diff reports no unrelated additions or removals$/,run:ctx=>{const d=ctx.report as PackageDiffReport;expect(d.added).toEqual([]);expect(d.removed).toEqual([]);expect(d.changed).toEqual(['[Content_Types].xml','word/document.xml']);}},
];
