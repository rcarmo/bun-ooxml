#!/usr/bin/python3
"""Optional development-only SmartArt import/export oracle using system UNO bindings."""
import json, os, signal, subprocess, sys, time, uuid
import uno, unohelper
from com.sun.star.task import XInteractionHandler
from com.sun.star.document.MacroExecMode import NEVER_EXECUTE
ROOT=os.path.abspath(sys.argv[1] if len(sys.argv)>1 else os.path.join(os.path.dirname(__file__),'../../artifacts/smartart-uno'))
OUT=ROOT+'/libreoffice'
os.makedirs(OUT,exist_ok=True)
report={'status':'running','producer':subprocess.check_output(['libreoffice','--version'],text=True).strip(),'interactions':[],'documents':{}}
class Handler(unohelper.Base,XInteractionHandler):
 def handle(self,request):
  report['interactions'].append(str(request.getRequest()))
  for continuation in request.getContinuations():
   if 'XInteractionAbort' in str(continuation): continuation.select();return
  raise RuntimeError('Unexpected interaction')
def prop(name,value):
 p=uno.createUnoStruct('com.sun.star.beans.PropertyValue');p.Name=name;p.Value=value;return p
def inspect(shape):
 row={'name':shape.Name,'type':shape.getShapeType(),'position':{'x':shape.Position.X,'y':shape.Position.Y},'size':{'width':shape.Size.Width,'height':shape.Size.Height}}
 try:row['text']=shape.getString()
 except Exception:pass
 if shape.supportsService('com.sun.star.drawing.GroupShape'):row['children']=[inspect(shape.getByIndex(i)) for i in range(shape.getCount())]
 return row
def snapshot(doc):
 pages=doc.getDrawPages();return [[inspect(pages.getByIndex(i).getByIndex(j)) for j in range(pages.getByIndex(i).getCount())] for i in range(pages.getCount())]
def walk(rows):
 for row in rows:
  yield row
  yield from walk(row.get('children',[]))
pipe='smartart_'+uuid.uuid4().hex
server=subprocess.Popen(['libreoffice','-env:UserInstallation=file://'+OUT+'/profile','--headless','--nologo','--nodefault','--nofirststartwizard','--accept=pipe,name='+pipe+';urp;StarOffice.ServiceManager'],stdout=open(OUT+'/server.stdout.log','w'),stderr=open(OUT+'/server.stderr.log','w'),start_new_session=True)
desktop=None
try:
 ctx=uno.getComponentContext();resolver=ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',ctx)
 for attempt in range(80):
  try:remote=resolver.resolve('uno:pipe,name='+pipe+';urp;StarOffice.ComponentContext');break
  except Exception:
   if server.poll() is not None:raise RuntimeError('server exited')
   time.sleep(.25)
 else:raise RuntimeError('connection timeout')
 desktop=remote.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',remote);handler=Handler()
 for filename in ['source.pptx','cross-presentation-copy.pptx','same-slide-copy.pptx']:
  doc=desktop.loadComponentFromURL(uno.systemPathToFileUrl(ROOT+'/'+filename),'_blank',0,(prop('Hidden',True),prop('ReadOnly',False),prop('MacroExecutionMode',NEVER_EXECUTE),prop('UpdateDocMode',0),prop('InteractionHandler',handler)))
  assert doc is not None
  try:
   before=snapshot(doc);leaf=[s for page in before for s in walk(page)];groups=[s for s in leaf if s['type'].endswith('GroupShape')];texts=[s['text'] for s in leaf if s.get('text')];assert len(before)==1
   expected_groups=2 if filename=='same-slide-copy.pptx' else 1
   assert len(groups)==expected_groups and all(len(g['children'])==6 for g in groups)
   report['documents'][filename]={'original':before,'groups':len(groups),'texts':texts}
   pdf=OUT+'/'+filename.replace('.pptx','.pdf');doc.storeToURL(uno.systemPathToFileUrl(pdf),(prop('FilterName','impress_pdf_Export'),prop('Overwrite',True)))
   saved=OUT+'/'+filename;doc.storeToURL(uno.systemPathToFileUrl(saved),(prop('FilterName','Impress MS PowerPoint 2007 XML'),prop('Overwrite',True)))
   reopened=desktop.loadComponentFromURL(uno.systemPathToFileUrl(saved),'_blank',0,(prop('Hidden',True),prop('MacroExecutionMode',NEVER_EXECUTE),prop('InteractionHandler',handler)))
   try:
    after=snapshot(reopened);report['documents'][filename]['reopened']=after
    next_groups=[s for page in after for s in walk(page) if s['type'].endswith('GroupShape')]
    assert len(next_groups)==expected_groups and all(len(g['children'])==6 for g in next_groups)
    maximum=0
    for a,b in zip(groups,next_groups):
     for kind in ['position','size']:
      for key in a[kind]:maximum=max(maximum,abs(a[kind][key]-b[kind][key]))
    assert maximum<=1
    report['documents'][filename]['maximumGroupGeometryDeltaHundredthsMm']=maximum
   finally:reopened.close(True)
  finally:doc.close(True)
 report['status']='passed'
 report['limits']=['No Microsoft PowerPoint application test','No SmartArt semantic editing or visual equivalence','LibreOffice imports these drawings as groups; six children and group geometry inspected','PPTX export classification-label content type prevents SDK reopening']
 assert not report['interactions']
except Exception as error:
 report['status']='failed';report['error']=str(error);raise
finally:
 if desktop:
  try:desktop.terminate()
  except Exception:pass
 try:server.wait(timeout=5)
 except subprocess.TimeoutExpired:os.killpg(server.pid,signal.SIGKILL)
 with open(OUT+'/report.json','w') as f:json.dump(report,f,indent=2)
 print(json.dumps({k:{'groups':v['groups'],'texts':v['texts']} for k,v in report['documents'].items()},indent=2))
