#!/usr/bin/python3
"""Development-only UNO oracle; Python is required by LibreOffice's UNO binding."""
import json, os, subprocess, sys, time, traceback, uuid
import uno
import unohelper
from com.sun.star.task import XInteractionHandler
from com.sun.star.document.MacroExecMode import NEVER_EXECUTE

ROOT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), '../../artifacts/graphics-uno'))
OUT = os.path.join(ROOT, 'uno')
os.makedirs(OUT, exist_ok=True)
PIPE = 'bun_graphics_' + uuid.uuid4().hex
report = {'status': 'running', 'producer': subprocess.check_output(['libreoffice', '--version'], text=True).strip() + ' UNO', 'transport': 'local named pipe', 'microsoftOffice': 'not tested', 'interactions': [], 'checks': {}, 'documents': {}}

class Interactions(unohelper.Base, XInteractionHandler):
    def handle(self, interaction):
        request = interaction.getRequest()
        report['interactions'].append(str(request))
        for continuation in interaction.getContinuations():
            if hasattr(continuation, 'abort'):
                continuation.abort()
                return
            if 'XInteractionAbort' in str(continuation):
                continuation.select()
                return
        raise RuntimeError('Unexpected interactive load request: ' + str(request))

handler = Interactions()
def prop(name, value):
    p = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    p.Name, p.Value = name, value
    return p

def scalar(value):
    if value is None or isinstance(value, (str, bool, int, float)):
        return value
    if hasattr(value, 'Top') and hasattr(value, 'Bottom') and hasattr(value, 'Left'):
        return {k: scalar(getattr(value, k)) for k in ['Top', 'Bottom', 'Left', 'Right']}
    if hasattr(value, 'StartColor') and hasattr(value, 'EndColor'):
        return {k: scalar(getattr(value, k)) for k in ['Style', 'StartColor', 'EndColor', 'Angle', 'Border', 'XOffset', 'YOffset', 'StartIntensity', 'EndIntensity', 'StepCount', 'ColorStops'] if hasattr(value, k)}
    if hasattr(value, 'Name') and hasattr(value, 'Value') and not hasattr(value, 'getShapeType'):
        return {'Name': value.Name, 'Value': scalar(value.Value)}
    if hasattr(value, 'StopOffset') and hasattr(value, 'StopColor'):
        return {'offset': value.StopOffset, 'color': scalar(value.StopColor)}
    if hasattr(value, 'Red') and hasattr(value, 'Green'):
        return {k: getattr(value, k) for k in ['Red', 'Green', 'Blue']}
    if hasattr(value, 'typeName') and hasattr(value, 'value'):
        return {'enum': value.typeName, 'value': str(value.value)}
    if isinstance(value, (tuple, list)):
        return [scalar(v) for v in value]
    if hasattr(value, 'getShapeType'):
        return {'name': value.Name, 'type': value.getShapeType()}
    if hasattr(value, '__dict__'):
        return {k: scalar(v) for k, v in value.__dict__.items() if not k.startswith('_')}
    return str(value)

PROPERTIES = ['RotateAngle', 'FillStyle', 'FillColor', 'FillTransparence', 'FillGradient', 'LineStyle', 'LineWidth', 'LineColor', 'LineStartName', 'LineEndName', 'LineJoint', 'LineCap', 'GraphicCrop', 'Transparency', 'StartShape', 'EndShape', 'StartGluePointIndex', 'EndGluePointIndex', 'CustomShapeGeometry']
def inspect_shape(shape):
    position, size = shape.getPosition(), shape.getSize()
    row = {'name': shape.Name, 'type': shape.getShapeType(), 'position': {'x': position.X, 'y': position.Y}, 'size': {'width': size.Width, 'height': size.Height}}
    try: row['text'] = shape.getString()
    except Exception: pass
    props = shape.getPropertySetInfo()
    row['properties'] = {k: scalar(shape.getPropertyValue(k)) for k in PROPERTIES if props.hasPropertyByName(k)}
    if shape.supportsService('com.sun.star.drawing.GroupShape'):
        row['children'] = [inspect_shape(shape.getByIndex(i)) for i in range(shape.getCount())]
    return row

def inspect(doc):
    pages = doc.getDrawPages()
    return [[inspect_shape(pages.getByIndex(i).getByIndex(j)) for j in range(pages.getByIndex(i).getCount())] for i in range(pages.getCount())]

def walk(shapes):
    for shape in shapes:
        yield shape
        yield from walk(shape.get('children', []))

def real_shapes(doc):
    def descend(shape):
        yield shape
        if shape.supportsService('com.sun.star.drawing.GroupShape'):
            for j in range(shape.getCount()): yield from descend(shape.getByIndex(j))
    for i in range(doc.getDrawPages().getCount()):
        page = doc.getDrawPages().getByIndex(i)
        for j in range(page.getCount()): yield from descend(page.getByIndex(j))

def metrics(data):
    leaves = [s for page in data for s in walk(page)]
    return {'pages': len(data), 'shapes': len(leaves), 'groups': sum(s['type'].endswith('GroupShape') for s in leaves), 'connectors': sum(s['type'].endswith('ConnectorShape') for s in leaves), 'graphics': sum(s['type'].endswith('GraphicObjectShape') for s in leaves), 'text': [s['text'] for s in leaves if s.get('text')], 'attachments': [{'name': s['name'], 'start': s['properties'].get('StartShape'), 'end': s['properties'].get('EndShape'), 'startSite': s['properties'].get('StartGluePointIndex'), 'endSite': s['properties'].get('EndGluePointIndex')} for s in leaves if s['type'].endswith('ConnectorShape')]}

server = subprocess.Popen(['libreoffice', '-env:UserInstallation=file://' + OUT + '/profile', '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--accept=pipe,name=' + PIPE + ';urp;StarOffice.ServiceManager'], stdout=open(OUT + '/server.stdout.log', 'w'), stderr=open(OUT + '/server.stderr.log', 'w'), start_new_session=True)
desktop = None
doc = None
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    context = None
    for attempt in range(80):
        try:
            context = resolver.resolve('uno:pipe,name=' + PIPE + ';urp;StarOffice.ComponentContext')
            break
        except Exception:
            if server.poll() is not None: raise RuntimeError('LibreOffice server exited')
            time.sleep(0.25)
    if context is None: raise RuntimeError('UNO connect timeout')
    desktop = context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)
    def load(path):
        result = desktop.loadComponentFromURL(uno.systemPathToFileUrl(path), '_blank', 0, (prop('Hidden', True), prop('ReadOnly', False), prop('MacroExecutionMode', NEVER_EXECUTE), prop('UpdateDocMode', 0), prop('InteractionHandler', handler)))
        if result is None: raise RuntimeError('Load refused: ' + path)
        return result
    doc = load(ROOT + '/graphics-showcase.pptx')
    original = inspect(doc)
    report['documents']['imported'] = original
    report['checks']['importedMetrics'] = metrics(original)
    imported_metrics = metrics(original)
    assert imported_metrics['pages'] == 4 and imported_metrics['groups'] == 1
    assert imported_metrics['connectors'] == 2 and imported_metrics['graphics'] == 4
    assert all(a['start'] and a['end'] for a in imported_metrics['attachments'])
    imported_shapes = [s for page in original for s in walk(page)]
    picture = next(s for s in imported_shapes if s['name'] == 'Picture 3')
    assert picture['properties']['Transparency'] == 20
    assert picture['properties']['RotateAngle'] == 35000
    assert picture['properties']['GraphicCrop'] == {'Top': 508, 'Bottom': 508, 'Left': 1016, 'Right': 1016}
    ellipse = next(s for s in imported_shapes if s.get('text') == 'Ellipse')
    assert ellipse['properties']['FillTransparence'] == 35
    gradient = next(s for s in imported_shapes if s.get('text') == 'Round rectangle')['properties']['FillGradient']
    assert gradient['StartColor'] == 0x60A5FA and gradient['EndColor'] == 0xDBEAFE
    report['checks']['literalImportProperties'] = 'crop, rotation, picture/shape transparency, gradient colours, group and connector attachments passed'
    edited_text = []
    moved = []
    for shape in real_shapes(doc):
        try: text = shape.getString()
        except Exception: text = ''
        if text in ['Input', 'Grouped A']:
            shape.setString(text + ' (UNO edited)')
            edited_text.append(text)
        if text == 'Round rectangle':
            position = shape.getPosition()
            position.X += 100
            shape.setPosition(position)
            moved.append(shape.Name)
    assert sorted(edited_text) == ['Grouped A', 'Input']
    assert len(moved) == 1
    edited = inspect(doc)
    report['documents']['edited'] = edited
    report['checks']['editedMetrics'] = metrics(edited)
    report['edits'] = {'text': edited_text, 'translatedShapes': moved, 'translation': '100 hundredths of mm along X (1 mm)'}
    expected = metrics(edited)
    for name, filter_name in [('graphics-uno-edited.odp', 'impress8'), ('graphics-uno-edited.pptx', 'Impress MS PowerPoint 2007 XML')]:
        path = OUT + '/' + name
        doc.storeToURL(uno.systemPathToFileUrl(path), (prop('FilterName', filter_name), prop('Overwrite', True), prop('InteractionHandler', handler)))
        reopened = load(path)
        try:
            observed = inspect(reopened)
            report['documents'][name] = observed
            actual = metrics(observed)
            checks = {'pagesPreserved': actual['pages'] == expected['pages'], 'shapeCountPreserved': actual['shapes'] == expected['shapes'], 'groupCountPreserved': actual['groups'] == expected['groups'], 'connectorCountPreserved': actual['connectors'] == expected['connectors'], 'graphicsCountPreserved': actual['graphics'] == expected['graphics'], 'editedLabelsPreserved': all(v in actual['text'] for v in ['Input (UNO edited)', 'Grouped A (UNO edited)']), 'connectorAttachmentsPresent': all(v['start'] and v['end'] for v in actual['attachments']), 'metrics': actual}
            report['checks'][name] = checks
            assert all(checks[k] for k in ['pagesPreserved', 'shapeCountPreserved', 'groupCountPreserved', 'connectorCountPreserved', 'graphicsCountPreserved', 'editedLabelsPreserved', 'connectorAttachmentsPresent'])
            pdf = OUT + '/' + name.rsplit('.', 1)[0] + '-' + name.rsplit('.', 1)[1] + '.pdf'
            reopened.storeToURL(uno.systemPathToFileUrl(pdf), (prop('FilterName', 'impress_pdf_Export'), prop('Overwrite', True)))
        finally: reopened.close(True)
    def compare_pages(before, after):
        maximum = 0
        errors = []
        scalar_keys = ['RotateAngle', 'FillStyle', 'FillColor', 'FillTransparence', 'LineWidth', 'LineColor', 'LineStartName', 'LineEndName', 'LineJoint', 'LineCap', 'Transparency', 'GraphicCrop', 'FillGradient']
        for index, (old_page, new_page) in enumerate(zip(before, after)):
            old_shapes, new_shapes = list(walk(old_page)), list(walk(new_page))
            if len(old_shapes) != len(new_shapes): errors.append('Page %s shape count' % index)
            for ordinal, (old, new) in enumerate(zip(old_shapes, new_shapes)):
                identity = '%s/%s:%s' % (index, ordinal, old['name'])
                for key in ['type', 'text']:
                    if old.get(key) != new.get(key): errors.append(identity + ' changed ' + key)
                for container in ['position', 'size']:
                    for key, value in old[container].items(): maximum = max(maximum, abs(value - new[container][key]))
                for key in scalar_keys:
                    if old['properties'].get(key) != new['properties'].get(key): errors.append(identity + ' changed ' + key)
                for endpoint in ['StartShape', 'EndShape', 'StartGluePointIndex', 'EndGluePointIndex']:
                    if old['properties'].get(endpoint) != new['properties'].get(endpoint): errors.append(identity + ' changed ' + endpoint)
        return {'maximumGeometryDeltaHundredthsMm': maximum, 'toleranceHundredthsMm': 1, 'scalarErrors': errors, 'passed': maximum <= 1 and not errors}
    for name in ['graphics-uno-edited.odp', 'graphics-uno-edited.pptx']:
        report['checks'][name]['scalarReadback'] = compare_pages(edited, report['documents'][name])
        assert report['checks'][name]['scalarReadback']['passed'], name + ' scalar/geometry preservation failed'
    report['checks']['noInteractiveRepairRequest'] = len(report['interactions']) == 0
    assert report['checks']['noInteractiveRepairRequest']
    report['limits'] = ['No Microsoft PowerPoint application test', 'No valid Office-produced SmartArt input', 'No visual fidelity or full preset/path/flip grammar comparison', 'UNO conversion renames title placeholders; geometry tolerance is 0.01 mm']
    report['status'] = 'passed'
except Exception as error:
    report['status'] = 'failed'
    report['error'] = str(error)
    report['traceback'] = traceback.format_exc()
    raise
finally:
    if doc:
        try: doc.close(True)
        except Exception: pass
    if desktop:
        try: desktop.terminate()
        except Exception: pass
    try: server.wait(timeout=5)
    except subprocess.TimeoutExpired:
        os.killpg(server.pid, 15)
        try: server.wait(timeout=3)
        except subprocess.TimeoutExpired: os.killpg(server.pid, 9)
    with open(OUT + '/report.json', 'w') as output: json.dump(report, output, indent=2, ensure_ascii=False)
    print(json.dumps({'status': report['status'], 'checks': report['checks']}, indent=2, ensure_ascii=False))
