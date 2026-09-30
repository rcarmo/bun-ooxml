import {test,expect} from 'bun:test';
import {OpcPackage,OoxmlError} from '../../src/index.ts';
import {writeZip} from '../../src/opc/zip.ts';
function source(){const e=(s:string)=>new TextEncoder().encode(s);return writeZip(new Map([
 ['[Content_Types].xml',e('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
 ['_rels/.rels',e('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
 ['word/document.xml',e('<document>Alpha</document>')],
 ]));}
test('deferred execution refuses before callback while immediate returns an opaque token unevaluated',async()=>{
 const bytes=source(),before=bytes.slice(),p=await OpcPackage.open(bytes);let ran=false,calls=0;
 expect(()=>p.transactionWithMode('deferred',()=>{ran=true;p.set('word/document.xml','<document>Beta</document>');})).toThrow(expect.objectContaining({code:'opc-deferred-transaction'}));
 expect(ran).toBe(false);expect(p.toBytes()).toEqual(before);expect(bytes).toEqual(before);
 const token={evaluate(){calls++;throw Error('not evaluated');},then(){calls++;throw Error('not awaited');}};
 const returned=p.transactionWithMode('immediate',()=>{p.set('word/document.xml','<document>Beta</document>');return token;});
 expect(returned).toBe(token);expect(calls).toBe(0);const saved=p.toBytes(),reopened=await OpcPackage.open(saved);expect(reopened.text('word/document.xml')).toBe('<document>Beta</document>');
 for(const name of ['[Content_Types].xml','_rels/.rels'])expect(reopened.get(name)).toEqual((await OpcPackage.open(before)).get(name));
 expect(bytes).toEqual(before);
});
test('portable immediate transaction rolls back callback and validation failure and retains async guard',async()=>{
 const p=await OpcPackage.open(source()),before=p.toBytes();
 expect(()=>p.transactionWithMode('immediate',()=>{p.set('word/document.xml','<document>Beta</document>');throw Error('callback');})).toThrow('callback');expect(p.toBytes()).toEqual(before);
 expect(()=>p.transactionWithMode('immediate',()=>{p.delete('word/document.xml');return {}; })).toThrow(expect.objectContaining({code:'opc-relationship-target-missing'}));expect(p.toBytes()).toEqual(before);
 let ran=false;expect(()=>p.transactionWithMode('immediate',async()=>{ran=true;})).toThrow(expect.objectContaining({code:'opc-async-transaction'}));expect(ran).toBe(false);expect(p.toBytes()).toEqual(before);
 expect(()=>p.transactionWithMode('invalid' as any,()=>{})).toThrow(OoxmlError);expect(p.toBytes()).toEqual(before);
});
