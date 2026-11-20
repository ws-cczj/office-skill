"""Read-only snapshots and optional geometry diagnostics from saved VSDX files."""
import hashlib
from pathlib import Path
import zipfile
import xml.etree.ElementTree as E

V = 'http://schemas.microsoft.com/office/visio/2012/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'v': V}
E.register_namespace('', V)
E.register_namespace('r', R)


def snapshot(filename, recursive=True):
    with zipfile.ZipFile(filename) as z:
        pages = E.fromstring(z.read('visio/pages/pages.xml'))
        rels = E.fromstring(z.read('visio/pages/_rels/pages.xml.rels'))
        targets = {r.get('Id'):r.get('Target') for r in rels}
        result = []
        for page in pages:
            rel = page.find('v:Rel', NS)
            name = targets[rel.get('{'+R+'}id')]
            part = 'visio/pages/' + name
            root = E.fromstring(z.read(part))
            shapes = root.findall('.//v:Shape', NS) if recursive else root.findall('v:Shapes/v:Shape', NS)
            items = []
            for s in shapes:
                cells = {c.get('N'):c.get('V') for c in s.findall('v:Cell', NS)}
                t = s.find('v:Text', NS)
                items.append({'id':int(s.get('ID')), 'key':s.get('NameU'), 'type':s.get('Type'),
                              'text': ''.join(t.itertext()) if t is not None else '', 'cells':cells})
            sheet = page.find('v:PageSheet', NS)
            result.append({'name':page.get('NameU'), 'part':part, 'shape_count':len(items),
                           'top_level_count':len(root.findall('v:Shapes/v:Shape', NS)),
                           'page_cells':{c.get('N'):c.get('V') for c in sheet.findall('v:Cell', NS)},
                           'shapes':items})
        return {'document':str(Path(filename).resolve()), 'state_source':'saved VSDX, not live unsaved state',
                'recursive':recursive, 'pages':result, 'quality_verdict':'not_assessed'}


def extract(filename):
    with zipfile.ZipFile(filename) as z:
        if any(n.startswith(('visio/media/', 'visio/masters/')) for n in z.namelist()):
            raise ValueError('Scene reconstruction currently requires standalone geometry without masters/media')
        summary = snapshot(filename)
        document = E.fromstring(z.read('visio/document.xml'))
        # Copy only formatting resources; do not carry source window state or package.
        support = [E.tostring(e,encoding='unicode') for e in document if e.tag.rsplit('}',1)[-1] in
                   {'DocumentSettings','Colors','FaceNames','StyleSheets','DocumentSheet'}]
        pages_root = E.fromstring(z.read('visio/pages/pages.xml'))
        model = {'format':'visio-native-scene-v1', 'source_sha256':hashlib.sha256(Path(filename).read_bytes()).hexdigest(),
                 'support':support, 'pages':[]}
        for entry, page in zip(summary['pages'], pages_root):
            content = E.fromstring(z.read(entry['part']))
            if any('Master' in s.attrib for s in content.findall('.//v:Shape',NS)):
                raise ValueError('Unresolved master references')
            if content.findall('.//v:ForeignData',NS):
                raise ValueError('Foreign data cannot be reconstructed as native geometry')
            model['pages'].append({'name':entry['name'], 'sheet':E.tostring(page.find('v:PageSheet',NS),encoding='unicode'),
                                   'content':E.tostring(content,encoding='unicode')})
        return model
